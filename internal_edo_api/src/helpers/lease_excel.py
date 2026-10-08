"""
Разбор Excel для пакетного создания договоров с проверкой каждой строки.

В отличие от parse_excel (agreemeant_parser), не падает на первой ошибке: каждая строка листа
«Договора» получает либо данные договора, либо список ошибок. Ошибки листа «Объект»
касаются всех договоров сразу, поэтому они фатальные — файл не принимается целиком.
"""

from __future__ import annotations

import re
from dataclasses import dataclass, field
from datetime import date, datetime
from io import BytesIO

import openpyxl
from pydantic import ValidationError

from helpers.agreemeant_parser import (
    BANK_BY_BIK,
    MONTHS_KZ,
    MONTHS_RU,
    ORG_TYPE_KZ,
    ROOM_TYPE_RU,
    _address_kz,
    _address_ru,
    _bank,
    _fmt,
    _genitive,
    _kbe,
    _org_names,
    _parse_object_sheet,
    _short_name,
    _words,
)
from core.phone import normalize_phone
from core.tin import tin_error
from schemas import ORG_TYPES_WITH_BIN, ContractDate, FinancialTerms, SubleaseContractData, TenantInfo

OBJECT_SHEET = "Объект"
CONTRACTS_SHEET = "Договора"
FIRST_DATA_ROW = 5  # строки 1-3 — заголовки, 4 — пример

_IBAN_RE = re.compile(r"^KZ\d{2}[A-Z0-9]{16}$")
_MIN_YEAR, _MAX_YEAR = 2000, 2100


@dataclass
class RowResult:
    row_number: int
    data: SubleaseContractData | None = None
    errors: list[str] = field(default_factory=list)


@dataclass
class ParsedWorkbook:
    category: str | None = None  # LEASE / SUBLEASE — «Тип договора» на листе «Объект»
    rows: list[RowResult] = field(default_factory=list)
    fatal: list[str] = field(default_factory=list)  # файл не принимается целиком


CATEGORY_BY_TITLE = {"аренда": "LEASE", "субаренда": "SUBLEASE"}


# ── Разбор значений ячеек ─────────────────────────────────────────────────────

class _Row:
    """Обёртка над строкой Excel: читает ячейки и копит ошибки, не прерывая разбор."""

    def __init__(self, values: tuple):
        self._values = values
        self.errors: list[str] = []

    def raw(self, i: int):
        return self._values[i] if i < len(self._values) else None

    def text(self, i: int, name: str, required: bool = True) -> str:
        value = self.raw(i)
        if isinstance(value, float) and value.is_integer():
            value = int(value)
        result = "" if value is None else str(value).strip()
        if required and not result:
            self.errors.append(f"{name}: не заполнено")
        return result

    def tin(self, i: int, name: str) -> str:
        value = self.raw(i)
        if isinstance(value, (int, float)):
            # Excel хранит длинные числа с потерей точности (920730401856 → 9,2E+11 → 920730000000),
            # поэтому ИИН, введённый числом, нельзя считать надёжным
            self.errors.append(f"{name}: введён числом — Excel мог исказить цифры, укажите ИИН как текст")
            return ""
        result = self.text(i, name)
        error = tin_error(result) if result else None
        if error:
            self.errors.append(f"{name} «{result}»: {error}")
        return result

    def digits(self, i: int, name: str, length: int, required: bool = True) -> str:
        """Номер документа из цифр. Excel срезает ведущие нули у чисел — восстанавливаем до длины."""
        value = self.raw(i)
        if isinstance(value, (int, float)) and float(value).is_integer():
            return str(int(value)).zfill(length)
        return self.text(i, name, required)

    def date(self, i: int, name: str) -> date | None:
        value = self.raw(i)
        if value is None or (isinstance(value, str) and not value.strip()):
            self.errors.append(f"{name}: не заполнено")
            return None
        if isinstance(value, datetime):
            d = value.date()
        elif isinstance(value, date):
            d = value
        else:
            d = None
            for fmt in ("%d.%m.%Y", "%Y-%m-%d"):
                try:
                    d = datetime.strptime(str(value).strip(), fmt).date()
                    break
                except ValueError:
                    continue
            if d is None:
                self.errors.append(f"{name}: неверный формат «{value}», нужно дд.мм.гггг")
                return None
        if not _MIN_YEAR <= d.year <= _MAX_YEAR:
            self.errors.append(f"{name}: неправдоподобный год {d.year}")
            return None
        return d

    def number(self, i: int, name: str, allow_zero: bool = False) -> int | None:
        value = self.raw(i)
        if isinstance(value, str):
            value = value.replace(" ", "").replace("\xa0", "").replace(",", ".")
        if value is None or value == "":
            self.errors.append(f"{name}: не заполнено")
            return None
        try:
            number = float(value)
        except (TypeError, ValueError):
            self.errors.append(f"{name}: должно быть числом (сейчас «{value}»)")
            return None
        if not number.is_integer():
            self.errors.append(f"{name}: должно быть целым числом (сейчас {value})")
            return None
        if number < 0 or (number == 0 and not allow_zero):
            self.errors.append(f"{name}: должно быть больше нуля")
            return None
        return int(number)


def _contract_date(d: date) -> ContractDate:
    return ContractDate(
        day=f"{d.day:02d}",
        month_kz=MONTHS_KZ[d.month],
        month_ru=MONTHS_RU[d.month],
        month_num=f"{d.month:02d}",
        year=str(d.year),
    )


def _validation_messages(e: ValidationError) -> list[str]:
    messages = []
    for err in e.errors():
        loc = [str(p) for p in err["loc"]]
        if loc and loc[-1] in ("iin", "director_iin"):
            messages.append(err["msg"].removeprefix("Value error, "))
        elif loc and loc[-1] == "org_type":
            messages.append(f"Тип орг. «{err.get('input')}»: допустимо ИП, ТОО, АО, ЖСШ")
        elif loc and loc[-1] == "phone":
            messages.append(f"Телефон: {err['msg'].removeprefix('Value error, ')}")
        else:
            messages.append(f"{'.'.join(loc)}: {err['msg']}")
    return messages


# ── Листы ─────────────────────────────────────────────────────────────────────

def _object_value(ws, key: str) -> str:
    for row in ws.iter_rows(min_row=3, values_only=True):
        if row and row[0] and str(row[0]).strip() == key:
            return str(row[1]).strip() if len(row) > 1 and row[1] is not None else ""
    return ""


def _parse_object(ws, result: ParsedWorkbook):
    kind = _object_value(ws, "Тип договора")
    # В старых файлах поля нет — это были только договоры субаренды
    result.category = CATEGORY_BY_TITLE.get(kind.lower(), None) if kind else "SUBLEASE"
    if result.category is None:
        result.fatal.append(f"Лист «{OBJECT_SHEET}»: «Тип договора» = «{kind}», допустимо Аренда или Субаренда")
        return None, None
    try:
        building, lessor = _parse_object_sheet(ws)
    except ValidationError as e:
        result.fatal += [f"Лист «{OBJECT_SHEET}»: {m}" for m in _validation_messages(e)]
        return None, None
    except Exception as e:
        result.fatal.append(f"Лист «{OBJECT_SHEET}»: {e}")
        return None, None

    required = {
        "Название здания (RU)": building.building_name_ru,
        "Адрес здания (RU)": building.building_address_ru,
        "Субарендодатель — Название": lessor.name_ru,
        "Субарендодатель — ФИО (KZ)": lessor.director_kz,
        "Субарендодатель — Счёт IBAN": lessor.account,
    }
    if result.category == "SUBLEASE":
        required["Основной договор аренды — номер (для субаренды)"] = building.lease_contract_num
        required["Основной договор аренды — дата (для субаренды)"] = building.lease_contract_date
    if lessor.org_type == "ИП":
        required["Арендодатель — Талон (для ИП)"] = lessor.talon
    for key, value in required.items():
        if not value or value.endswith("«»"):
            result.fatal.append(f"Лист «{OBJECT_SHEET}»: не заполнено «{key}»")
    if lessor.bik.strip().upper() not in BANK_BY_BIK:
        result.fatal.append(f"Лист «{OBJECT_SHEET}»: БИК арендодателя «{lessor.bik}» нет в справочнике банков")
    return building, lessor


def _parse_contract_row(values: tuple, row_number: int, building, lessor, category: str) -> RowResult:
    r = _Row(values)

    contract_number = r.text(0, "Номер договора")
    signed = r.date(1, "Дата подписания")
    start = r.date(2, "Дата начала аренды")
    end = r.date(3, "Дата окончания аренды")
    floor = r.text(4, "Этаж", required=False)
    room_type = r.text(5, "Тип помещения")
    area = r.number(6, "Площадь")
    rent_sqm = r.number(7, "Ставка за м²")
    communal = r.number(8, "Коммунальные", allow_zero=True)

    org_type = r.text(9, "Тип орг. арендатора").upper()
    org_name = r.text(10, "Название арендатора")
    full_name = r.text(11, "ФИО арендатора")
    tin = r.tin(12, "ИИН/БИН арендатора")
    id_num = r.digits(13, "Уд. личн. арендатора", length=9, required=False)
    id_date = r.text(14, "Дата уд. арендатора", required=False)
    talon = r.text(15, "Талон арендатора", required=False)
    addr_kz = r.text(16, "Адрес арендатора (KZ)")
    addr_ru = r.text(17, "Адрес арендатора (RU)")
    bik = r.text(18, "БИК арендатора").upper()
    account = r.text(19, "Счёт IBAN арендатора").upper().replace(" ", "")
    try:
        phone = normalize_phone(r.raw(20))
    except ValueError as e:
        r.errors.append(f"Телефон арендатора: {e}")
        phone = ""

    # ИИН директора — только у ТОО/АО (для ИП колонка отключена: его ИИН и есть ИИН/БИН из колонки M)
    director_iin = r.tin(21, "ИИН директора арендатора") if org_type in ORG_TYPES_WITH_BIN else ""

    if start and end and end <= start:
        r.errors.append("Дата окончания аренды должна быть позже даты начала")
    if org_type and org_type not in ORG_TYPE_KZ:
        r.errors.append(f"Тип орг. арендатора: «{org_type}», допустимо {', '.join(ORG_TYPE_KZ)}")
    if room_type and room_type not in ROOM_TYPE_RU:
        r.errors.append(f"Тип помещения: «{room_type}», допустимо {', '.join(ROOM_TYPE_RU)}")
    if bik and bik not in BANK_BY_BIK:
        r.errors.append(f"БИК арендатора: «{bik}» нет в справочнике банков")
    if account and not _IBAN_RE.match(account):
        r.errors.append(f"Счёт IBAN арендатора: неверный формат «{account}» (KZ + 18 символов)")
    if id_date:
        id_date_value = r.raw(14)
        if isinstance(id_date_value, datetime):
            id_date = id_date_value.strftime("%d.%m.%Y")

    if r.errors:
        return RowResult(row_number, errors=r.errors)

    rent_total = area * rent_sqm
    rent_sqm_kz, rent_sqm_ru = _words(rent_sqm)
    rent_total_kz, rent_total_ru = _words(rent_total)
    communal_kz, communal_ru = _words(communal)
    name_kz, name_ru = _org_names(org_type, org_name)

    try:
        data = SubleaseContractData(
            category=category,
            contract_number=contract_number,
            signed_date=_contract_date(signed),
            start_date=_contract_date(start),
            end_date=_contract_date(end),
            financial=FinancialTerms(
                floor=floor,
                room_type=room_type,
                area=str(area),
                rent_per_sqm=_fmt(rent_sqm),
                rent_per_sqm_words_kz=rent_sqm_kz,
                rent_per_sqm_words_ru=rent_sqm_ru,
                rent_total=_fmt(rent_total),
                rent_total_words_kz=rent_total_kz,
                rent_total_words_ru=rent_total_ru,
                communal=_fmt(communal),
                communal_words_kz=communal_kz,
                communal_words_ru=communal_ru,
            ),
            building=building,
            lessor=lessor,
            tenant=TenantInfo(
                org_type=org_type,
                director_iin=director_iin,
                name_kz=name_kz,
                name_ru=name_ru,
                director_kz=full_name,
                director_ru=_genitive(full_name),
                director_short=_short_name(full_name),
                iin=tin,
                talon=talon,
                id_num=id_num,
                id_date=id_date,
                address_kz=_address_kz(addr_kz),
                address_ru=_address_ru(addr_ru),
                bik=bik,
                bank=_bank(bik),
                kbe=_kbe(org_type),
                account=account,
                phone=phone,
            ),
        )
    except ValidationError as e:
        return RowResult(row_number, errors=_validation_messages(e))
    return RowResult(row_number, data=data)


# ── Главная функция ───────────────────────────────────────────────────────────

def parse_lease_workbook(content: bytes) -> ParsedWorkbook:
    result = ParsedWorkbook()
    try:
        wb = openpyxl.load_workbook(BytesIO(content), data_only=True)
    except Exception as e:
        result.fatal.append(f"Не удалось открыть Excel: {e}")
        return result

    for sheet in (OBJECT_SHEET, CONTRACTS_SHEET):
        if sheet not in wb.sheetnames:
            result.fatal.append(f"Лист «{sheet}» не найден")
    if result.fatal:
        return result

    building, lessor = _parse_object(wb[OBJECT_SHEET], result)
    if result.fatal:
        return result

    seen: dict[str, int] = {}
    for row_number, values in enumerate(
        wb[CONTRACTS_SHEET].iter_rows(min_row=FIRST_DATA_ROW, values_only=True), start=FIRST_DATA_ROW
    ):
        if not any(v is not None and str(v).strip() for v in values):
            continue
        row = _parse_contract_row(values, row_number, building, lessor, result.category)
        if row.data is not None:
            number = row.data.contract_number
            if number in seen:
                row = RowResult(row_number, errors=[f"Номер договора «{number}» уже есть в строке {seen[number]}"])
            else:
                seen[number] = row_number
        result.rows.append(row)

    if not result.rows:
        result.fatal.append(f"На листе «{CONTRACTS_SHEET}» нет договоров (заполняются с {FIRST_DATA_ROW}-й строки)")
    return result
