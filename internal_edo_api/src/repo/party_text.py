"""
Тексты сторон договора, которые зависят от формы стороны (ИП или ТОО/АО), а не от типа договора:
преамбула («…действующего на основании Устава / талона…») и блок реквизитов.

Формулировки взяты из договоров заказчика: ТОО — из договора аренды №1/129, ИП — из субаренды №116.
Шаблоны .docx содержат только {{parties_ru}}, {{lessor_details_ru}} и т.п., поэтому форма
арендодателя и арендатора может быть любой в обоих типах договора.
"""

from __future__ import annotations

from docxtpl import Listing

from core.banks import bank_by_bik

from schemas import ORG_TYPES_WITH_BIN, SubleaseContractData
from schemas.service_agreement import LessorInfo, TenantInfo

ROLES = {
    # category: (арендодатель RU, арендатор RU, арендодатель KZ, арендатор KZ)
    "LEASE": ("Арендодатель", "Арендатор", "Жалға беруші", "Жалға алушы"),
    "SUBLEASE": ("Субарендодатель", "Субарендатор", "Қосалқы жалға беруші", "Қосалқы жалға алушы"),
}


def _bank_ru(party: LessorInfo | TenantInfo) -> str:
    bank = bank_by_bik(party.bik)
    return bank.ru if bank else party.bank


def _bank_kz(party: LessorInfo | TenantInfo) -> str:
    bank = bank_by_bik(party.bik)
    return bank.kz if bank else party.bank


def _director_ru(party: LessorInfo | TenantInfo) -> str:
    return party.director_ru.rstrip(", ")  # родительный падеж из Excel приходит с запятой на конце


def _intro_ru(party: LessorInfo | TenantInfo, role: str) -> str:
    if party.org_type in ORG_TYPES_WITH_BIN:
        director_iin = f" ИИН:{party.director_iin}" if party.director_iin else ""
        return (f"{party.name_ru}, БИН {party.iin} именуемое в дальнейшем «{role}», "
                f"в лице директора {_director_ru(party)}{director_iin}, действующего на основании Устава")
    return (f"{party.name_ru}, именуемый в дальнейшем «{role}», в лице директора {_director_ru(party)} "
            f"ИИН:{party.iin}, действующего на основании талона №{party.talon}")


def _intro_kz(party: LessorInfo | TenantInfo, role: str) -> str:
    if party.org_type in ORG_TYPES_WITH_BIN:
        director_iin = f" ЖСН:{party.director_iin}" if party.director_iin else ""
        return (f'{party.name_kz}, БСН {party.iin} бұдан әрі "{role}" деп аталатын, '
                f"Жарғы негізінде әрекет ететін директоры {party.director_kz}{director_iin}")
    return (f"{party.name_kz}, директоры {party.director_kz} ЖСН:{party.iin}, "
            f'№{party.talon} талон негізінде әрекет ететін, бұдан әрі "{role}" деп аталатын')


def _details_ru(party: LessorInfo | TenantInfo, data: SubleaseContractData, is_lessor: bool) -> list[str]:
    b = data.building
    if party.org_type in ORG_TYPES_WITH_BIN:
        lines = [f"БИН {party.iin}", f"Юридический адрес: {party.address_ru}"]
        if is_lessor:
            lines.append(f"Фактический адрес {b.building_name_ru}: {b.building_address_ru}")
    else:
        lines = [f"ИИН: {party.iin}", f"Адрес: {party.address_ru}"]
    if party.id_num:
        lines.append(f"Уд. лич. №{party.id_num} от {party.id_date} г.")
    if party.org_type not in ORG_TYPES_WITH_BIN and party.talon:
        lines.append(f"Талон №{party.talon}")
    lines += [f"Банк: {_bank_ru(party)}", f"КБе: {party.kbe}", f"БИК: {party.bik}"]
    lines.append(f"ИИК: {party.account}" if party.org_type in ORG_TYPES_WITH_BIN else f"Номер счета: {party.account}")
    return lines


def _details_kz(party: LessorInfo | TenantInfo, data: SubleaseContractData, is_lessor: bool) -> list[str]:
    b = data.building
    if party.org_type in ORG_TYPES_WITH_BIN:
        lines = [f"БСН {party.iin}", f"Заңды мекен-жайы: {party.address_kz}"]
        if is_lessor:
            lines.append(f"Нақты мекен-жайы {b.building_name_kz}: {b.building_address_kz}")
    else:
        lines = [f"ЖСН: {party.iin}", f"Мекен-жайы: {party.address_kz}"]
    if party.id_num:
        lines.append(f"Жеке куәлік: №{party.id_num}, {party.id_date} ж.")
    if party.org_type not in ORG_TYPES_WITH_BIN and party.talon:
        lines.append(f"Талон №{party.talon}")
    lines += [f"Банк: {_bank_kz(party)}", f"КБе: {party.kbe}", f"БИК: {party.bik}"]
    lines.append(f"ИИК: {party.account}" if party.org_type in ORG_TYPES_WITH_BIN else f"Шот нөмірі: {party.account}")
    return lines


def party_context(data: SubleaseContractData, phone_display: str) -> dict:
    """Переменные шаблона, зависящие от формы сторон."""
    lessor_role_ru, tenant_role_ru, lessor_role_kz, tenant_role_kz = ROLES[data.category]
    lo, t = data.lessor, data.tenant

    tenant_ru = _details_ru(t, data, is_lessor=False)
    tenant_kz = _details_kz(t, data, is_lessor=False)
    if phone_display:
        tenant_ru.append(f"Тел: {phone_display}")
        tenant_kz.append(f"Тел: {phone_display}")

    return {
        "parties_ru": f"{_intro_ru(lo, lessor_role_ru)} с одной стороны и {_intro_ru(t, tenant_role_ru)}",
        "parties_kz": f"{_intro_kz(lo, lessor_role_kz)}, бір тараптан және {_intro_kz(t, tenant_role_kz)}",
        # Listing: строки через \n становятся переносами строк внутри абзаца
        "lessor_details_ru": Listing("\n".join(_details_ru(lo, data, is_lessor=True))),
        "lessor_details_kz": Listing("\n".join(_details_kz(lo, data, is_lessor=True))),
        "tenant_details_ru": Listing("\n".join(tenant_ru)),
        "tenant_details_kz": Listing("\n".join(tenant_kz)),
    }
