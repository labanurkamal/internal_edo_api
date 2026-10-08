"""
ИИН/БИН Казахстана: 12 цифр, последняя — контрольная.

Контрольная цифра = сумма(цифра_i × вес_i) mod 11 по первым 11 цифрам, веса 1..11.
Если получилось 10 — пересчёт с весами 3..11,1,2; снова 10 — номер недействителен.
Проверка ловит опечатки до создания договора: иначе договор создастся, а подпись
отклонится как SIGNING_ERROR_TIN (ИИН/БИН из сертификата не совпадёт с участником).
"""

from __future__ import annotations

_WEIGHTS_1 = tuple(range(1, 12))
_WEIGHTS_2 = (3, 4, 5, 6, 7, 8, 9, 10, 11, 1, 2)


def tin_control_digit(first11: str) -> int | None:
    digits = [int(c) for c in first11]
    s = sum(d * w for d, w in zip(digits, _WEIGHTS_1)) % 11
    if s == 10:
        s = sum(d * w for d, w in zip(digits, _WEIGHTS_2)) % 11
        if s == 10:
            return None
    return s


def tin_error(tin: str) -> str | None:
    """Текст ошибки или None, если ИИН/БИН корректен."""
    if len(tin) != 12 or not tin.isdigit():
        return "должен состоять из 12 цифр"
    control = tin_control_digit(tin[:11])
    if control is None or control != int(tin[11]):
        return "неверная контрольная цифра — проверьте номер"
    return None


def validate_tin(tin: str) -> str:
    """Для pydantic: возвращает номер или поднимает ValueError с понятным текстом."""
    error = tin_error(tin)
    if error:
        raise ValueError(f"ИИН/БИН «{tin}»: {error}")
    return tin
