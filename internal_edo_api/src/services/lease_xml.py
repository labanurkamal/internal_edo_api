"""
XML договора для подписи ЭЦП (NCALayer).
Структура как у актов и актов сверки: стороны во вложенных <seller>/<buyer>, номер и дата
в document_number/document_date. Подписывается именно этот XML, а целостность PDF
обеспечивает поле <hash> (SHA-256 PDF).
"""

from __future__ import annotations

import xml.etree.ElementTree as ET
from typing import Any


_DECLARATION = b'<?xml version="1.0" encoding="UTF-8"?>\n'


def _append(parent: ET.Element, key: str, value: Any) -> None:
    node = ET.SubElement(parent, key)
    if isinstance(value, dict):
        for k, v in value.items():
            _append(node, k, v)
    else:
        node.text = "" if value is None else str(value)


def build_lease_xml(fields: dict[str, Any]) -> bytes:
    """Строит XML из словаря: dict → вложенный тег, None → пустой тег. Порядок полей сохраняется."""
    root = ET.Element("lease_contract")
    for key, value in fields.items():
        _append(root, key, value)
    # Декларация с двойными кавычками: consumer-parks (btodict) заменяет ' на " во всём теле сообщения
    # и ломается на <?xml version='1.0'?>, которую по умолчанию пишет ElementTree
    return _DECLARATION + ET.tostring(root, encoding="utf-8", xml_declaration=False)
