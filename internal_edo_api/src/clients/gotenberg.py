"""
Клиент Gotenberg: конвертация .docx → .pdf через LibreOffice.
Единственная ответственность — HTTP-вызов конвертера. Не знает про договоры и MinIO.
"""

from __future__ import annotations

import logging

import httpx

from core.config import settings

logger = logging.getLogger(__name__)

DOCX_MEDIA_TYPE = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
_CONVERT_ENDPOINT = "/forms/libreoffice/convert"
_PDF_MAGIC = b"%PDF-"
_ZIP_MAGIC = b"PK\x03\x04"


class PdfConversionError(Exception):
    """Gotenberg недоступен или вернул не PDF."""


class GotenbergClient:
    def __init__(self, base_url: str | None = None, timeout_seconds: float | None = None):
        self._base_url = (base_url or settings.gotenberg.url).rstrip("/")
        self._timeout = httpx.Timeout(timeout_seconds or settings.gotenberg.timeout_seconds, connect=5.0)

    async def docx_to_pdf(self, docx: bytes, filename: str = "document.docx") -> bytes:
        """Отправляет .docx в Gotenberg, возвращает байты PDF."""
        # Gotenberg конвертирует любые байты (как обычный текст), поэтому проверяем вход сами:
        # .docx — это zip-архив
        if not docx.startswith(_ZIP_MAGIC):
            raise PdfConversionError(f"{filename}: входные данные не являются .docx")

        files = {"files": (filename, docx, DOCX_MEDIA_TYPE)}
        try:
            async with httpx.AsyncClient(base_url=self._base_url, timeout=self._timeout) as client:
                response = await client.post(_CONVERT_ENDPOINT, files=files)
        except httpx.HTTPError as e:
            raise PdfConversionError(f"Gotenberg недоступен ({self._base_url}): {e!r}") from e

        if response.status_code != 200:
            raise PdfConversionError(
                f"Gotenberg вернул {response.status_code}: {response.text[:300]}"
            )

        pdf = response.content
        if not pdf.startswith(_PDF_MAGIC):
            raise PdfConversionError("Gotenberg вернул ответ, который не является PDF")

        logger.info("docx → pdf: %s (%d → %d байт)", filename, len(docx), len(pdf))
        return pdf
