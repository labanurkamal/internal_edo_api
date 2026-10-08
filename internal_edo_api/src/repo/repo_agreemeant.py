"""
Рендеринг договоров через docxtpl.
Добавить новый тип договора = добавить новый класс-наследник, бизнес-логика не трогается.
"""

from __future__ import annotations
from abc import ABC, abstractmethod
from io import BytesIO
from pathlib import Path

from dataclasses import dataclass
from functools import lru_cache

from docxtpl import DocxTemplate

from schemas import SubleaseContractData
from repo.party_text import party_context


class ContractRenderer(ABC):
    """Интерфейс рендерера. Каждый тип договора — своя реализация."""

    @abstractmethod
    def render(self, data) -> BytesIO:
        """Рендерит договор, возвращает .docx в памяти."""
        ...

    @abstractmethod
    def filename(self, data) -> str:
        """Имя файла для сохранения в MinIO."""
        ...


class DocxtplRenderer(ContractRenderer):
    """Базовый рендерер через docxtpl. Принимает любой шаблон и контекст."""

    def __init__(self, template_path: str | Path):
        self._template_path = Path(template_path)
        if not self._template_path.exists():
            raise FileNotFoundError(f"Шаблон не найден: {self._template_path}")

    def _render_context(self, context: dict) -> BytesIO:
        tpl = DocxTemplate(str(self._template_path))
        tpl.render(context)
        buf = BytesIO()
        tpl.save(buf)
        buf.seek(0)
        return buf

    def render(self, data) -> BytesIO:
        raise NotImplementedError("Используй конкретный рендерер")

    def filename(self, data) -> str:
        raise NotImplementedError("Используй конкретный рендерер")


class SubleaseContractRenderer(DocxtplRenderer):
    """Рендерер договора субаренды."""

    def render(self, data: SubleaseContractData) -> BytesIO:
        context = data.to_template_context()
        # Преамбула и реквизиты сторон зависят от формы (ИП / ТОО), а не от шаблона
        context.update(party_context(data, phone_display=context["tenant_phone"]))
        return self._render_context(context)

    def filename(self, data: SubleaseContractData) -> str:
        import re
        num = data.contract_number.replace("/", "_")
        # Убираем тип орг и кавычки: ТОО «Алтын Нур» → Алтын_Нур
        name = re.sub(r"^(ИП|ТОО|АО|ЖСШ)\s*[«\"']?", "", data.tenant.name_ru)
        name = name.replace("»", "").replace("«", "").strip().replace(" ", "_")
        kind = "Аренды" if data.category == "LEASE" else "Субаренды"
        return f"Договор_{kind}_{num}_{name}.docx"


# ──────────────────────────────────────────────
# Шаблоны договоров: отдельный .docx на каждый тип, данные и конвейер общие.
# Новая редакция шаблона = новый файл с новой версией; созданные договоры хранят,
# какой версией они сделаны (lease_contracts.template_code / template_version).
# ──────────────────────────────────────────────
TEMPLATES_DIR = Path(__file__).parent.parent / "api" / "templates"


@dataclass(frozen=True)
class ContractTemplate:
    code: str      # = category
    version: int
    path: Path


CONTRACT_TEMPLATES: dict[str, ContractTemplate] = {
    "LEASE": ContractTemplate("LEASE", 1, TEMPLATES_DIR / "Договор_Аренды_v1.docx"),
    "SUBLEASE": ContractTemplate("SUBLEASE", 3, TEMPLATES_DIR / "Договор_Субаренды_v3.docx"),
}


@lru_cache
def renderer_for(category: str) -> SubleaseContractRenderer:
    return SubleaseContractRenderer(CONTRACT_TEMPLATES[category].path)
