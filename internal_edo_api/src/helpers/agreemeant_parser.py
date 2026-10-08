"""
Парсер Excel → list[SubleaseContractData].
Читает два листа: 'Объект' (один раз) и 'Договора' (каждая строка = договор).
"""

from __future__ import annotations
from datetime import datetime
from io import BytesIO

import openpyxl
from num2words import num2words

from core.banks import BANKS
from core.declension import genitive_fio

from schemas import (
    SubleaseContractData, ContractDate, FinancialTerms,
    BuildingInfo, LessorInfo, TenantInfo,
)

# ── Справочники ───────────────────────────────────────────────────────────────

MONTHS_KZ = {
    1: "қаңтар",   2: "ақпан",    3: "наурыз",   4: "сәуір",
    5: "мамыр",    6: "маусым",   7: "шілде",    8: "тамыз",
    9: "қыркүйек", 10: "қазан",   11: "қараша",  12: "желтоқсан",
}
MONTHS_RU = {
    1: "января",   2: "февраля",  3: "марта",    4: "апреля",
    5: "мая",      6: "июня",     7: "июля",     8: "августа",
    9: "сентября", 10: "октября", 11: "ноября",  12: "декабря",
}

# Тип орг RU → KZ
ORG_TYPE_KZ = {"ИП": "ЖК", "ТОО": "ЖШС", "АО": "АҚ", "ЖСШ": "ЖСШ"}

# КБе по типу орг
KBE_MAP = {"ИП": "19", "ТОО": "17", "АО": "17", "ЖСШ": "17"}

# Банк по БИК
# Банк по БИК (казахское название; русское — в core.banks). Неизвестный БИК — ошибка строки.
BANK_BY_BIK = {bik: bank.kz for bik, bank in BANKS.items()}

# Тип помещения KZ → RU (для выпадающего списка в Excel)
ROOM_TYPE_RU = {
    "бутик":    "бутик",
    "офис":     "офис",
    "қойма":    "склад",
    "павильон": "павильон",
    "дүкен":    "магазин",
}


# ── Вспомогательные функции ───────────────────────────────────────────────────

def _s(val) -> str:
    return str(val).strip() if val is not None else ""


def _int(val) -> int:
    return int(float(_s(val).replace(" ", "").replace("\xa0", "")))


def _parse_date(raw) -> ContractDate:
    d = None
    if isinstance(raw, datetime):
        d = raw
    else:
        for fmt in ("%d.%m.%Y", "%Y-%m-%d %H:%M:%S", "%Y-%m-%d"):
            try:
                d = datetime.strptime(raw, fmt)
            except ValueError:
                continue

        if d is None:
            raise ValueError(f"Қолдау көрсетілмейтін күн форматы: {raw!r}")
    return ContractDate(
        day=f"{d.day:02d}",
        month_kz=MONTHS_KZ[d.month],
        month_ru=MONTHS_RU[d.month],
        month_num=f"{d.month:02d}",
        year=str(d.year),
    )


def _words(amount: int) -> tuple[str, str]:
    return num2words(amount, lang="kz"), num2words(amount, lang="ru")


def _fmt(amount: int) -> str:
    return f"{amount:,}".replace(",", " ")


def _short_name(full: str) -> str:
    parts = full.strip().split()
    if len(parts) >= 3:
        return f"{parts[0]} {parts[1][0]}.{parts[2][0]}."
    if len(parts) == 2:
        return f"{parts[0]} {parts[1][0]}."
    return full


def _org_names(org_type: str, name: str) -> tuple[str, str]:
    org = org_type.strip().upper()
    kz  = ORG_TYPE_KZ.get(org, org)
    return f"{kz} «{name}»", f"{org} «{name}»"


def _genitive(full: str) -> str:
    """Родительный падеж ФИО для русского текста («в лице директора …»)."""
    return genitive_fio(full)


def _address_ru(street: str) -> str:
    s = street.strip()
    return s if s.startswith("ул.") else f"ул. {s}"


def _address_kz(street: str) -> str:
    s = street.strip()
    return s if s.endswith("үй") else f"{s} үй"


def _bank(bik: str) -> str:
    return BANK_BY_BIK.get(bik.strip().upper(), f"БИК: {bik}")


def _kbe(org_type: str) -> str:
    return KBE_MAP.get(org_type.strip().upper(), "19")


# ── Парсер листа "Объект" ─────────────────────────────────────────────────────

# Новое название поля в шаблоне Excel → старое (ключ) для совместимости со старыми файлами
OBJECT_KEY_ALIASES = {
    "Субарендодатель — Тип орг.": ("Арендодатель — Тип орг.",),
    "Субарендодатель — Название": ("Арендодатель — Название",),
    "Субарендодатель — ФИО (KZ)": ("Арендодатель — ФИО директора (KZ)",),
    "Субарендодатель — ФИО (RU)": ("Арендодатель — ФИО директора (RU, род. падеж)",),
    "Субарендодатель — ИИН": ("Арендодатель — ИИН/БИН",),
    "Субарендодатель — ИИН директора": ("Арендодатель — ИИН директора (для ТОО/АО)",),
    "Субарендодатель — Талон": ("Арендодатель — Талон (для ИП)",),
    "Субарендодатель — Уд. личн.": ("Арендодатель — Уд. личн.",),
    "Субарендодатель — Дата уд.": ("Арендодатель — Дата уд.",),
    "Субарендодатель — Адрес (KZ)": ("Арендодатель — Адрес (KZ)",),
    "Субарендодатель — Адрес (RU)": ("Арендодатель — Адрес (RU)",),
    "Субарендодатель — БИК": ("Арендодатель — БИК",),
    "Субарендодатель — Счёт IBAN": ("Арендодатель — Счёт IBAN",),
    "Договор аренды номер": ("Основной договор аренды — номер (для субаренды)",),
    "Договор аренды дата": ("Основной договор аренды — дата (для субаренды)",),
}

def _parse_object_sheet(ws) -> tuple[BuildingInfo, LessorInfo]:
    """
    Лист 'Объект' — вертикальный формат: колонка A = ключ, колонка B = значение.
    Порядок строк фиксирован (см. create_excel_template.py).
    """
    rows = {}
    for row in ws.iter_rows(min_row=3, values_only=True):
        key = _s(row[0]) if row[0] else ""
        val = _s(row[1]) if len(row) > 1 and row[1] else ""
        if key:
            rows[key] = val

    def g(key: str) -> str:
        # Шаблон Excel переименовал поля («Субарендодатель — …» → «Арендодатель — …»);
        # старые файлы тоже принимаются
        for name in (key, *OBJECT_KEY_ALIASES.get(key, ())):
            if rows.get(name):
                return rows[name]
        return ""

    # Арендодатель
    lessor_org_type = g("Субарендодатель — Тип орг.")
    lessor_org_name = g("Субарендодатель — Название")
    lessor_name_kz, lessor_name_ru = _org_names(lessor_org_type, lessor_org_name)
    lessor_full_kz  = g("Субарендодатель — ФИО (KZ)")
    lessor_full_ru  = g("Субарендодатель — ФИО (RU)") or _genitive(lessor_full_kz)
    lessor_bik      = g("Субарендодатель — БИК")

    lessor = LessorInfo(
        org_type=lessor_org_type.strip().upper() or "ИП",
        director_iin=g("Субарендодатель — ИИН директора"),
        name_kz=lessor_name_kz,
        name_ru=lessor_name_ru,
        director_kz=lessor_full_kz,
        director_ru=lessor_full_ru,
        director_short=_short_name(lessor_full_kz),
        iin=g("Субарендодатель — ИИН"),
        talon=g("Субарендодатель — Талон"),
        id_num=g("Субарендодатель — Уд. личн."),
        id_date=g("Субарендодатель — Дата уд."),
        address_kz=g("Субарендодатель — Адрес (KZ)"),
        address_ru=g("Субарендодатель — Адрес (RU)"),
        bik=lessor_bik,
        bank=_bank(lessor_bik),
        kbe=_kbe(lessor_org_type),
        account=g("Субарендодатель — Счёт IBAN"),
    )

    # Дата госакта — "16 января 2004" → KZ: "2004 жылғы 16 қаңтардағы"
    act_date_raw = g("Госакт дата")  # "16 января 2004"
    act_date_kz  = act_date_raw  # пользователь вводит RU, KZ вычисляем
    try:
        parts = act_date_raw.split()  # ["16", "января", "2004"]
        ru_to_kz = {
            "января":"қаңтар","февраля":"ақпан","марта":"наурыз","апреля":"сәуір",
            "мая":"мамыр","июня":"маусым","июля":"шілде","августа":"тамыз",
            "сентября":"қыркүйек","октября":"қазан","ноября":"қараша","декабря":"желтоқсан"
        }
        act_date_kz = f"{parts[2]} жылғы {parts[0]} {ru_to_kz.get(parts[1], parts[1])}дағы"
    except Exception:
        act_date_kz = act_date_raw

    building = BuildingInfo(
        building_name_kz=g("Название здания (KZ)"),
        building_name_ru=g("Название здания (RU)"),
        room_type_kz=g("Тип помещения (KZ)"),
        room_type_ru=g("Тип помещения (RU)"),
        building_address_kz=g("Адрес здания (KZ)"),
        building_address_ru=g("Адрес здания (RU)"),
        owner_name_kz=g("Собственник (KZ)"),
        owner_name_ru=g("Собственник (RU)"),
        owner_act_num=g("Госакт номер"),
        owner_act_date_kz=act_date_kz,
        owner_act_date_ru=act_date_raw,
        lease_contract_num=g("Договор аренды номер"),
        lease_contract_date=g("Договор аренды дата"),
    )

    return building, lessor


# ── Парсер листа "Договора" ───────────────────────────────────────────────────

def _parse_contracts_sheet(
    ws,
    building: BuildingInfo,
    lessor: LessorInfo,
) -> list[SubleaseContractData]:
    """Строки с 5-й (строки 1-3 заголовки, строка 4 пример)."""
    results = []
    errors  = []

    for row_idx, row in enumerate(ws.iter_rows(min_row=5, values_only=True), start=5):
        if not any(row):
            continue

        def col(i: int) -> str:
            return _s(row[i]) if i < len(row) and row[i] is not None else ""

        try:
            contract_number = col(0)   # A
            if not contract_number:
                continue

            signed_date = _parse_date(col(1))  # B
            start_date  = _parse_date(col(2))  # C
            end_date    = _parse_date(col(3))  # D

            floor     = col(4)                  # E
            room_type = col(5)                  # F — бутик / офис / қойма
            area      = _int(col(6))            # G
            rent_sqm  = _int(col(7))            # H
            communal  = _int(col(8))            # I
            rent_total = area * rent_sqm

            rent_sqm_kz,   rent_sqm_ru   = _words(rent_sqm)
            rent_total_kz, rent_total_ru = _words(rent_total)
            communal_kz,   communal_ru   = _words(communal)

            financial = FinancialTerms(
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
            )

            # Арендатор (J-U)
            t_org_type  = col(9)    # J
            t_org_name  = col(10)   # K
            t_full_name = col(11)   # L
            t_iin       = col(12)   # M
            t_id_num    = col(13)   # N
            t_id_date   = col(14)   # O
            t_talon     = col(15)   # P
            t_addr_kz   = col(16)   # Q
            t_addr_ru   = col(17)   # R
            t_bik       = col(18)   # S
            t_account   = col(19)   # T
            t_phone     = col(20)   # U

            t_name_kz, t_name_ru = _org_names(t_org_type, t_org_name)

            tenant = TenantInfo(
                name_kz=t_name_kz,
                name_ru=t_name_ru,
                director_kz=t_full_name,
                director_ru=_genitive(t_full_name),
                director_short=_short_name(t_full_name),
                iin=t_iin,
                talon=t_talon,
                id_num=t_id_num,
                id_date=t_id_date,
                address_kz=_address_kz(t_addr_kz),
                address_ru=_address_ru(t_addr_ru),
                bik=t_bik,
                bank=_bank(t_bik),
                kbe=_kbe(t_org_type),
                account=t_account,
                phone=t_phone,
            )

            results.append(SubleaseContractData(
                contract_number=contract_number,
                signed_date=signed_date,
                start_date=start_date,
                end_date=end_date,
                financial=financial,
                building=building,
                lessor=lessor,
                tenant=tenant,
            ))

        except Exception as e:
            errors.append(f"Строка {row_idx}: {e}")

    if errors:
        raise ValueError("Ошибки при парсинге:\n" + "\n".join(errors))

    return results


# ── Главная функция ───────────────────────────────────────────────────────────

def parse_excel(file: BytesIO | str) -> list[SubleaseContractData]:
    """Читает Excel, возвращает список SubleaseContractData."""
    wb = openpyxl.load_workbook(file, data_only=True)

    if "Объект" not in wb.sheetnames:
        raise ValueError("Лист 'Объект' не найден в Excel файле")
    if "Договора" not in wb.sheetnames:
        raise ValueError("Лист 'Договора' не найден в Excel файле")

    building, lessor = _parse_object_sheet(wb["Объект"])
    return _parse_contracts_sheet(wb["Договора"], building, lessor)