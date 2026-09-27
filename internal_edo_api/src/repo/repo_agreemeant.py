"""
Рендеринг договоров через docxtpl.
Добавить новый тип договора = добавить новый класс-наследник, бизнес-логика не трогается.
"""

from __future__ import annotations
from abc import ABC, abstractmethod
from io import BytesIO
from pathlib import Path

from docxtpl import DocxTemplate

from schemas import SubleaseContractData


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
        return self._render_context(context)

    def filename(self, data: SubleaseContractData) -> str:
        import re
        num = data.contract_number.replace("/", "_")
        # Убираем тип орг и кавычки: ТОО «Алтын Нур» → Алтын_Нур
        name = re.sub(r"^(ИП|ТОО|АО|ЖСШ)\s*[«\"']?", "", data.tenant.name_ru)
        name = name.replace("»", "").replace("«", "").strip().replace(" ", "_")
        return f"Договор_Субаренды_{num}_{name}.docx"


# ──────────────────────────────────────────────
# Реестр рендереров. Новый тип → одна строка здесь.
# ──────────────────────────────────────────────
RENDERERS: dict[str, type[ContractRenderer]] = {
    "sublease": SubleaseContractRenderer,
    # "lease": LeaseContractRenderer,      # пример расширения
    # "employment": EmploymentRenderer,    # пример расширения
}