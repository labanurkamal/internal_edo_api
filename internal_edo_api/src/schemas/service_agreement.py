"""
Pydantic v2 модели для договоров субаренды.
Новый тип договора = новая модель + новый шаблон. Бизнес-логика не трогается.
"""

from __future__ import annotations
from pydantic import BaseModel, Field


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
    lease_contract_num:  str
    lease_contract_date: str


class LessorInfo(BaseModel):
    """Субарендодатель — из листа 'Объект', меняется редко."""
    name_kz:      str
    name_ru:      str
    director_kz:  str   # именительный: Сайланбаева А.Б.
    director_ru:  str   # родительный:  Сайланбаевой А.Б.,
    director_short: str # сокращённо:   Сайланбаева А.Б.
    iin:          str
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
    """Субарендатор — из листа 'Договора', каждый раз разный."""
    name_kz:        str
    name_ru:        str
    director_kz:    str   # именительный полный
    director_ru:    str   # родительный (авто или ручной)
    director_short: str   # авто: Фамилия И.О.
    iin:            str
    talon:          str
    id_num:         str
    id_date:        str
    address_kz:     str
    address_ru:     str
    bik:            str
    bank:           str   # авто по bik
    kbe:            str   # авто по типу орг
    account:        str
    phone:          str


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
    """Все данные для генерации одного договора субаренды."""
    contract_number: str
    signed_date:     ContractDate
    start_date:      ContractDate
    end_date:        ContractDate
    financial:       FinancialTerms
    building:        BuildingInfo
    lessor:          LessorInfo
    tenant:          TenantInfo

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
            "tenant_phone":             t.phone,
        }