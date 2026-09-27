"""
ContractService — оркестратор: рендеринг + сохранение в MinIO.
Единственная ответственность — координация. Не знает про HTTP и не знает про XML.
"""

from __future__ import annotations

import os.path
from dataclasses import dataclass

from core.config import settings
from schemas import SubleaseContractData
from repo import SubleaseContractRenderer, ContractRenderer
from db.minio import MinioStorage


@dataclass
class GeneratedContract:
    object_name: str
    download_url: str
    filename: str


class ContractService:
    """
    Принимает данные → рендерит .docx → конвертирует в PDF (если converter задан)
    → сохраняет в MinIO → возвращает presigned URL.

    Пример без конвертации (отдаёт .docx):
        service = ContractService(renderer, storage)

    Пример с конвертацией (отдаёт .pdf, открывается в Google Docs Viewer):
        service = ContractService(renderer, storage, converter=GotenbergConverter())
    """

    def __init__(
            self,
            renderer: ContractRenderer,
            storage: MinioStorage,
    ):
        self._renderer = renderer
        self._storage = storage

    def generate(self, data: SubleaseContractData) -> GeneratedContract:
        """Одиночная генерация."""
        docx_buf = self._renderer.render(data)
        docx_name = self._renderer.filename(data)

        final_buf = docx_buf
        final_name = docx_name
        content_type = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"

        object_name = f"subleases/{final_name}"
        self._storage.upload(object_name, final_buf, content_type=content_type)
        url = "https://docs.google.com/viewer?url=" + self._storage.presigned_url(object_name)

        return GeneratedContract(
            object_name=object_name,
            download_url=url,
            filename=final_name,
        )

    def generate_bulk(self, items: list[SubleaseContractData]) -> list[GeneratedContract]:
        """Массовая генерация."""
        return [self.generate(item) for item in items]