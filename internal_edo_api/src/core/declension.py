"""
Родительный падеж ФИО для русского текста договора: «в лице директора Ерошиной Елены Леонидовны».

Склонение — pytrovich (порт petrovich, правила для русских антропонимов). Пол определяем сами,
потому что казахские формы petrovich не знает: «…қызы/кызы» — женщина, «…ұлы/улы» — мужчина.
Если пол определить нельзя, ФИО возвращается без изменений: лучше несклонённое, чем склонённое неверно.
"""

from __future__ import annotations

from functools import lru_cache

from pytrovich.detector import PetrovichGenderDetector
from pytrovich.enums import Case, Gender, NamePart
from pytrovich.maker import PetrovichDeclinationMaker

_FEMALE_ENDINGS = ("қызы", "кызы", "овна", "евна", "ична", "инична")
_MALE_ENDINGS = ("ұлы", "улы", "ович", "евич", "ьич", "ич")
_FEMALE_SURNAMES = ("ова", "ева", "ёва", "ина", "ына", "ая")
_MALE_SURNAMES = ("ов", "ев", "ёв", "ин", "ын", "ий", "ой")
_PARTS = (NamePart.LASTNAME, NamePart.FIRSTNAME, NamePart.MIDDLENAME)


@lru_cache
def _maker() -> PetrovichDeclinationMaker:
    return PetrovichDeclinationMaker()


@lru_cache
def _detector() -> PetrovichGenderDetector:
    return PetrovichGenderDetector()


def _gender(parts: list[str]) -> Gender | None:
    for word in parts[1:]:
        w = word.lower()
        if w.endswith(_FEMALE_ENDINGS):
            return Gender.FEMALE
        if w.endswith(_MALE_ENDINGS):
            return Gender.MALE
    # Казахская форма «Ғалымжанқызы Гауһар»: признак пола стоит первым словом
    first = parts[0].lower()
    if first.endswith(("қызы", "кызы")):
        return Gender.FEMALE
    if first.endswith(("ұлы", "улы")):
        return Gender.MALE

    names = [p for p in parts[1:] if "." not in p]
    if names:
        try:
            detected = _detector().detect(firstname=names[0], middlename=names[1] if len(names) > 1 else None)
        except Exception:
            detected = None
        if detected in (Gender.MALE, Gender.FEMALE):
            return detected

    if first.endswith(_FEMALE_SURNAMES):
        return Gender.FEMALE
    if first.endswith(_MALE_SURNAMES):
        return Gender.MALE
    return None


def genitive_fio(full_name: str) -> str:
    """«Фамилия Имя Отчество» → родительный падеж. Инициалы («А.Б.») не склоняются."""
    parts = full_name.split()
    if not parts:
        return full_name
    gender = _gender(parts)
    if gender is None:
        return full_name

    result = []
    for word, kind in zip(parts, _PARTS):
        if "." in word or word.lower().endswith(("қызы", "кызы", "ұлы", "улы")):
            result.append(word)
            continue
        try:
            result.append(_maker().make(kind, gender, Case.GENITIVE, word))
        except Exception:
            result.append(word)
    result += parts[len(_PARTS):]
    return " ".join(result)
