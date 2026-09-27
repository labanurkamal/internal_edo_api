"""
FastAPI роутер. Тонкий слой — только HTTP.
"""

from __future__ import annotations
from pathlib import Path

from fastapi import APIRouter, HTTPException, Depends, UploadFile, File
from fastapi.responses import StreamingResponse
from pydantic import BaseModel
from io import BytesIO
import zipfile

from schemas import SubleaseContractData
from services import ContractService, GeneratedContract
from repo import SubleaseContractRenderer
from helpers import parse_excel
from db.minio import MinioStorage
from core.config import Settings

router = APIRouter(prefix="/contracts", tags=["contracts"])

TEMPLATE_PATH = Path(__file__).parent.parent / "templates" / "Договор_Субаренды_ШАБЛОН_v2.docx"


# ── Response schemas ──────────────────────────────────────────────────────────

class ContractResponse(BaseModel):
    filename: str
    download_url: str
    object_name: str


class BulkContractResponse(BaseModel):
    total: int
    contracts: list[ContractResponse]


# ── Dependency injection ──────────────────────────────────────────────────────

def get_service() -> ContractService:
    renderer = SubleaseContractRenderer(TEMPLATE_PATH)
    storage = MinioStorage(Settings())
    return ContractService(renderer, storage)


def _to_response(contract: GeneratedContract) -> ContractResponse:
    return ContractResponse(
        filename=contract.filename,
        download_url=contract.download_url,
        object_name=contract.object_name,
    )


# ── Endpoints ─────────────────────────────────────────────────────────────────

@router.post(
    "/subleases/generate",
    response_model=ContractResponse,
    summary="Сгенерировать один договор (JSON)",
)
async def generate_sublease(
    data: SubleaseContractData,
    service: ContractService = Depends(get_service),
) -> ContractResponse:
    try:
        return _to_response(service.generate(data))
    except FileNotFoundError as e:
        raise HTTPException(status_code=500, detail=str(e))
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Ошибка генерации: {e}")


@router.post(
    "/subleases/bulk/json",
    response_model=BulkContractResponse,
    summary="Массовая генерация (JSON-массив)",
)
async def generate_bulk_json(
    items: list[SubleaseContractData],
    service: ContractService = Depends(get_service),
) -> BulkContractResponse:
    if not items:
        raise HTTPException(status_code=400, detail="Список пустой")
    try:
        contracts = service.generate_bulk(items)
        return BulkContractResponse(total=len(contracts), contracts=[_to_response(c) for c in contracts])
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Ошибка: {e}")


@router.post(
    "/subleases/bulk/excel",
    response_model=BulkContractResponse,
    summary="Массовая генерация из Excel → ссылки MinIO",
)
async def generate_bulk_excel(
    file: UploadFile = File(..., description="Excel по шаблону Договора_Шаблон.xlsx"),
    service: ContractService = Depends(get_service),
) -> BulkContractResponse:
    if not file.filename.endswith((".xlsx", ".xlsm")):
        raise HTTPException(status_code=400, detail="Нужен файл .xlsx")
    try:
        print(1)
        items = parse_excel(BytesIO(await file.read()))
    except ValueError as e:
        raise HTTPException(status_code=422, detail=str(e))
    except Exception as e:
        raise HTTPException(status_code=400, detail=f"Не удалось прочитать Excel: {e}")

    if not items:
        raise HTTPException(status_code=400, detail="В файле нет данных (строки 5+)")

    try:
        contracts = service.generate_bulk(items)
        return BulkContractResponse(total=len(contracts), contracts=[_to_response(c) for c in contracts])
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Ошибка генерации: {e}")


@router.post(
    "/subleases/bulk/excel/zip",
    summary="Массовая генерация из Excel → ZIP-архив (без MinIO)",
    response_class=StreamingResponse,
)
async def generate_bulk_excel_zip(
    file: UploadFile = File(...),
) -> StreamingResponse:
    """Возвращает ZIP со всеми договорами прямо в ответе. Удобно для небольших пачек."""
    if not file.filename.endswith((".xlsx", ".xlsm")):
        raise HTTPException(status_code=400, detail="Нужен файл .xlsx")
    try:
        items = parse_excel(BytesIO(await file.read()))
    except ValueError as e:
        raise HTTPException(status_code=422, detail=str(e))

    if not items:
        raise HTTPException(status_code=400, detail="В файле нет данных")

    renderer = SubleaseContractRenderer(TEMPLATE_PATH)
    zip_buf = BytesIO()
    with zipfile.ZipFile(zip_buf, "w", zipfile.ZIP_DEFLATED) as zf:
        for item in items:
            doc_buf = renderer.render(item)
            zf.writestr(renderer.filename(item), doc_buf.read())

    zip_buf.seek(0)
    return StreamingResponse(
        zip_buf,
        media_type="application/zip",
        headers={"Content-Disposition": "attachment; filename=contracts.zip"},
    )


from urllib.parse import quote

@router.get(
    "/subleases/template",
    summary="Скачать Excel-шаблон для заполнения",
    response_class=StreamingResponse,
)
async def download_excel_template() -> StreamingResponse:
    """Отдаёт готовый Excel-шаблон."""
    import sys, os
    sys.path.insert(0, str(Path(__file__).parent.parent))
    from api.templates.agreemeant_dowload import create_template

    buf = BytesIO()
    create_template(buf)
    buf.seek(0)

    filename = "Договора_Шаблон.xlsx"
    # ASCII-фолбэк + UTF-8 нұсқасы (RFC 5987)
    ascii_fallback = "dogovora_shablon.xlsx"
    content_disposition = (
        f"attachment; filename=\"{ascii_fallback}\"; "
        f"filename*=UTF-8''{quote(filename)}"
    )

    return StreamingResponse(
        buf,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": content_disposition},
    )