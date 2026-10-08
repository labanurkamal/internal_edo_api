"""
Телефоны Казахстана.

Хранение (БД, XML, JSON): +77011112233 — без пробелов.
Вывод в договоре (.docx/.pdf): +7 701 111 22 33.
"""

from __future__ import annotations

import re

_NOISE_RE = re.compile(r"[\s\-()]")


def normalize_phone(raw: str | int | float | None) -> str:
    """
    Принимает +7…, 8… и 7… (11 цифр), в том числе с пробелами, дефисами и скобками.
    Возвращает +7XXXXXXXXXX; пустое значение → "". Неверный формат → ValueError.
    """
    if raw is None:
        return ""
    if isinstance(raw, float) and raw.is_integer():  # Excel хранит номер, введённый без +, как число
        raw = int(raw)
    value = _NOISE_RE.sub("", str(raw)).removesuffix(".0")  # "77074567890.0" из старого парсера Excel
    if not value:
        return ""

    digits = value[1:] if value.startswith("+") else value
    if not digits.isdigit() or len(digits) != 11:
        raise ValueError(f"«{raw}»: нужно 11 цифр, например +77011112233 или 87011112233")
    if value.startswith("+") and digits[0] != "7":
        raise ValueError(f"«{raw}»: после + должна идти 7")
    if digits[0] not in "78":
        raise ValueError(f"«{raw}»: должен начинаться с +7, 8 или 7")
    return "+7" + digits[1:]


def format_phone(phone: str | None) -> str:
    """+77011112233 → +7 701 111 22 33. Пустое или нестандартное значение возвращается как есть."""
    if not phone or not re.fullmatch(r"\+7\d{10}", phone):
        return phone or ""
    d = phone[2:]
    return f"+7 {d[0:3]} {d[3:6]} {d[6:8]} {d[8:10]}"
