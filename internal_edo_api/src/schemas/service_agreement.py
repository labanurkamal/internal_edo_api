"""
Pydantic v2 модели для договоров субаренды.
Новый тип договора = новая модель + новый шаблон. Бизнес-логика не трогается.
"""

from __future__ import annotations
from typing import Annotated, Literal

from pydantic import AfterValidator, BaseModel, BeforeValidator, Field, StringConstraints, model_validator

from core.phone import format_phone, normalize_phone
from core.tin import validate_tin

# ИИН/БИН: 12 цифр с верной контрольной цифрой. По нему сверяется подписант ЭЦП, поэтому
# ошибка здесь потом даёт отказ в подписании (SIGNING_ERROR_TIN_*)
Tin = Annotated[str, StringConstraints(strip_whitespace=True), AfterValidator(validate_tin)]
# ИИН директора: пусто (ИП, у которого ИИН уже в iin) или корректный ИИН
OptionalTin = Annotated[str, StringConstraints(strip_whitespace=True),
                        AfterValidator(lambda v: v if v == "" else validate_tin(v))]

# Телефон хранится как +77011112233; принимается +7…, 8…, 7… с пробелами/дефисами
Phone = Annotated[str, BeforeValidator(normalize_phone)]

# Форма стороны определяет текст преамбулы и реквизитов:
# ИП — ИИН и талон; ТОО/АО — БИН, Устав и ИИН директора
OrgType = Literal["ИП", "ТОО", "АО", "ЖСШ"]
ORG_TYPES_WITH_BIN = ("ТОО", "АО")


# ── Вспомогательные модели ────────────────────────────────────────────────────

class BilingualStr(BaseModel):
    """Строка на двух языках."""
    kz: str
    ru: str


class ContractDate(BaseModel):
    """Дата с месяцем на двух языках."""
    day:       str
    month_kz:  str
    month_ru:  str
    month_num: str
    year:      str


# ── Объект / Здание ───────────────────────────────────────────────────────────

class BuildingInfo(BaseModel):
    """Данные здания — из листа 'Объект', общие для всех договоров."""
    building_name_kz:    str
    building_name_ru:    str
    room_type_kz:        str   # тұрғын емес / тұрғын
    room_type_ru:        str   # нежилое / жилое
    building_address_kz: str
    building_address_ru: str
    owner_name_kz:       str
    owner_name_ru:       str
    owner_act_num:       str
    owner_act_date_kz:   str   # напр. 2004 жылғы 16 қаңтардағы
    owner_act_date_ru:   str   # напр. 16 января 2004г
    # Основной договор аренды — только для субаренды (на нём основано право сдавать в субаренду)
    lease_contract_num:  str = ""
    lease_contract_date: str = ""


class LessorInfo(BaseModel):
    """Арендодатель/субарендодатель — из листа 'Объект', меняется редко."""
    org_type:     OrgType = "ИП"
    name_kz:      str
    name_ru:      str
    director_kz:  str   # именительный: Сайланбаева А.Б.
    director_ru:  str   # родительный:  Сайланбаевой А.Б.,
    director_short: str # сокращённо:   Сайланбаева А.Б.
    iin:          Tin   # ИИН для ИП, БИН для ТОО/АО
    director_iin: OptionalTin = ""  # ИИН директора — для ТОО/АО
    talon:        str
    id_num:       str
    id_date:      str
    address_kz:   str
    address_ru:   str
    bik:          str
    bank:         str   # авто по bik
    kbe:          str   # авто по типу орг
    account:      str


# ── Арендатор ─────────────────────────────────────────────────────────────────

class TenantInfo(BaseModel):
    """Арендатор/субарендатор — из листа 'Договора', каждый раз разный."""
    org_type:       OrgType = "ИП"
    name_kz:        str
    name_ru:        str
    director_kz:    str   # именительный полный
    director_ru:    str   # родительный (авто или ручной)
    director_short: str   # авто: Фамилия И.О.
    iin:            Tin   # ИИН для ИП, БИН для ТОО/АО — по нему сверяется подписант ЭЦП
    director_iin:   OptionalTin = ""  # ИИН директора — обязателен для ТОО/АО
    talon:          str
    id_num:         str
    id_date:        str
    address_kz:     str
    address_ru:     str
    bik:            str
    bank:           str   # авто по bik
    kbe:            str   # авто по типу орг
    account:        str
    phone:          Phone = ""


# ── Финансы ───────────────────────────────────────────────────────────────────

class FinancialTerms(BaseModel):
    floor:               str
    room_type:           str   # бутик / офис / склад
    area:                str
    rent_per_sqm:        str
    rent_per_sqm_words_kz: str
    rent_per_sqm_words_ru: str
    rent_total:          str
    rent_total_words_kz: str
    rent_total_words_ru: str
    communal:            str
    communal_words_kz:   str
    communal_words_ru:   str


# ── Главная модель договора ───────────────────────────────────────────────────

class SubleaseContractData(BaseModel):
    """Все данные для генерации одного договора аренды или субаренды (category)."""
    category:        Literal["LEASE", "SUBLEASE"] = "SUBLEASE"
    contract_number: str
    signed_date:     ContractDate
    start_date:      ContractDate
    end_date:        ContractDate
    financial:       FinancialTerms
    building:        BuildingInfo
    lessor:          LessorInfo
    tenant:          TenantInfo

    @model_validator(mode="after")
    def _check_category_fields(self):
        if self.category == "SUBLEASE" and not (self.building.lease_contract_num and self.building.lease_contract_date):
            raise ValueError("для субаренды нужны номер и дата основного договора аренды")
        if self.tenant.org_type in ORG_TYPES_WITH_BIN and not self.tenant.director_iin:
            raise ValueError(f"для арендатора {self.tenant.org_type} нужен ИИН директора")
        return self

    def to_template_context(self) -> dict:
        """Разворачивает модель в плоский dict для docxtpl."""
        d  = self.signed_date
        s  = self.start_date
        e  = self.end_date
        f  = self.financial
        b  = self.building
        lo = self.lessor
        t  = self.tenant
        return {
            # Договор
            "contract_number":          self.contract_number,
            # Дата подписания
            "date_day":                 d.day,
            "date_month_kz":            d.month_kz,
            "date_month_ru":            d.month_ru,
            "date_year":                d.year,
            "date_month_num":           d.month_num,
            # Начало
            "start_day":                s.day,
            "start_month_kz":           s.month_kz,
            "start_month_ru":           s.month_ru,
            "start_month_num":          s.month_num,
            "start_year":               s.year,
            # Конец
            "end_day":                  e.day,
            "end_month_kz":             e.month_kz,
            "end_month_ru":             e.month_ru,
            "end_year":                 e.year,
            # Финансы
            "floor":                    f.floor,
            "room_type":                f.room_type,
            "area":                     f.area,
            "rent_per_sqm":             f.rent_per_sqm,
            "rent_per_sqm_words_kz":    f.rent_per_sqm_words_kz,
            "rent_per_sqm_words_ru":    f.rent_per_sqm_words_ru,
            "rent_total":               f.rent_total,
            "rent_total_words_kz":      f.rent_total_words_kz,
            "rent_total_words_ru":      f.rent_total_words_ru,
            "communal":                 f.communal,
            "communal_words_kz":        f.communal_words_kz,
            "communal_words_ru":        f.communal_words_ru,
            # Здание
            "building_name_kz":         b.building_name_kz,
            "building_name_ru":         b.building_name_ru,
            "room_type_kz":             b.room_type_kz,
            "room_type_ru":             b.room_type_ru,
            "building_address_kz":      b.building_address_kz,
            "building_address_ru":      b.building_address_ru,
            "owner_name_kz":            b.owner_name_kz,
            "owner_name_ru":            b.owner_name_ru,
            "owner_act_num":            b.owner_act_num,
            "owner_act_date_kz":        b.owner_act_date_kz,
            "owner_act_date_ru":        b.owner_act_date_ru,
            "lease_contract_num":       b.lease_contract_num,
            "lease_contract_date":      b.lease_contract_date,
            # Арендодатель
            "lessor_name_kz":           lo.name_kz,
            "lessor_name_ru":           lo.name_ru,
            "lessor_director_kz":       lo.director_kz,
            "lessor_director_ru":       lo.director_ru,
            "lessor_director_short":    lo.director_short,
            "lessor_iin":               lo.iin,
            "lessor_talon":             lo.talon,
            "lessor_id_num":            lo.id_num,
            "lessor_id_date":           lo.id_date,
            "lessor_address_kz":        lo.address_kz,
            "lessor_address_ru":        lo.address_ru,
            "lessor_bik":               lo.bik,
            "lessor_bank":              lo.bank,
            "lessor_kbe":               lo.kbe,
            "lessor_account":           lo.account,
            # Арендатор
            "tenant_name_kz":           t.name_kz,
            "tenant_name_ru":           t.name_ru,
            "tenant_director_kz":       t.director_kz,
            "tenant_director_ru":       t.director_ru,
            "tenant_director_short":    t.director_short,
            "tenant_iin":               t.iin,
            "tenant_talon":             t.talon,
            "tenant_id_num":            t.id_num,
            "tenant_id_date":           t.id_date,
            "tenant_address_kz":        t.address_kz,
            "tenant_address_ru":        t.address_ru,
            "tenant_bik":               t.bik,
            "tenant_bank":              t.bank,
            "tenant_kbe":               t.kbe,
            "tenant_account":           t.account,
            "tenant_phone":             format_phone(t.phone),  # в договоре: +7 701 111 22 33
        }


# Модель общая для аренды и субаренды; старое имя оставлено для совместимости
LeaseContractData = SubleaseContractData
