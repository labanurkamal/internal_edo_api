"""
Договоры аренды и субаренды для подписания в EDO. Тонкий слой — только HTTP.
Тип договора — в данных: category в JSON, «Тип договора» на листе «Объект» в Excel.
В отличие от /contracts/subleases/* создаёт записи в БД (documents, participants, lease_contracts),
PDF и XML для подписи.
"""

from __future__ import annotations

import uuid
from datetime import datetime
from functools import lru_cache
from urllib.parse import quote

from fastapi import APIRouter, Depends, File, HTTPException, Query, Response, UploadFile, status
from pydantic import BaseModel

from clients import GotenbergClient, PdfConversionError
from core.config import Settings
from db.minio import MinioStorage
from db.psql_engine import session as session_factory
from schemas import SubleaseContractData
from services.lease_batch import BatchFileRejected, LeaseBatchService
from services.lease_contract import (
    LeaseContractConflict,
    LeaseContractResult,
    LeaseContractService,
)

router = APIRouter(prefix="/edo/leases", tags=["edo-leases"])
_XLSX_TYPE = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"


# ── Response schemas ──────────────────────────────────────────────────────────

class LeaseContractResponse(BaseModel):
    id: int
    category: str
    contract_number: str
    status: str | None
    pdf_id: uuid.UUID
    xml_id: uuid.UUID
    pdf_url: str
    created: bool


class BatchAcceptedResponse(BaseModel):
    batch_id: uuid.UUID
    category: str  # LEASE / SUBLEASE — из «Тип договора» на листе «Объект»
    total: int
    accepted: int  # ушли в фоновую обработку
    rejected: int  # не прошли проверку — см. errors_url
    status_url: str  # прогресс фоновой обработки
    errors_url: str | None  # Excel со строками-ошибками; None — ошибок при проверке нет


class BatchRowErrorResponse(BaseModel):
    row_number: int
    status: str
    error: str | None


class BatchStatusResponse(BaseModel):
    batch_id: uuid.UUID
    filename: str | None
    created: datetime
    total: int
    counts: dict[str, int]  # rejected / pending / processing / done / error
    finished: bool
    errors: list[BatchRowErrorResponse]
    errors_url: str | None  # None — ошибок нет


# ── Dependency injection ──────────────────────────────────────────────────────

@lru_cache
def _storage() -> MinioStorage:
    return MinioStorage(Settings())


@lru_cache
def get_lease_service() -> LeaseContractService:
    return LeaseContractService(
        storage=_storage(),
        converter=GotenbergClient(),
        session_factory=session_factory,
    )


@lru_cache
def get_batch_service() -> LeaseBatchService:
    return LeaseBatchService(storage=_storage(), session_factory=session_factory, lease_service=get_lease_service())


def _status_url(batch_id: uuid.UUID) -> str:
    return router.url_path_for("get_batch_status", batch_id=str(batch_id))


def _errors_url(batch_id: uuid.UUID) -> str:
    return router.url_path_for("download_batch_errors", batch_id=str(batch_id))


def _to_response(result: LeaseContractResult) -> LeaseContractResponse:
    return LeaseContractResponse(**result.__dict__)


# ── Endpoints ─────────────────────────────────────────────────────────────────

@router.post(
    "",
    response_model=LeaseContractResponse,
    summary="Создать договор аренды или субаренды для подписания (JSON, тип — в поле category)",
    responses={
        200: {"description": "Договор уже существовал — возвращён без изменений (created=false)"},
        409: {"description": "Номер договора занят другим арендатором"},
        502: {"description": "Не удалось сконвертировать в PDF"},
    },
    status_code=status.HTTP_201_CREATED,
)
async def create_sublease(
    data: SubleaseContractData,
    response: Response,
    user_id: str | None = Query(None, title="Кто создал"),
    service: LeaseContractService = Depends(get_lease_service),
) -> LeaseContractResponse:
    try:
        result = await service.create(data, created_by=user_id)
    except LeaseContractConflict as e:
        raise HTTPException(status_code=status.HTTP_409_CONFLICT, detail=str(e))
    except PdfConversionError as e:
        raise HTTPException(status_code=status.HTTP_502_BAD_GATEWAY, detail=str(e))
    if not result.created:
        response.status_code = status.HTTP_200_OK
    return _to_response(result)


@router.post(
    "/bulk/excel",
    response_model=BatchAcceptedResponse,
    status_code=status.HTTP_202_ACCEPTED,
    summary="Пакетная загрузка из Excel: проверка сразу, создание в фоне",
    description=(
        "Строки, прошедшие проверку, создаются в фоне — прогресс в GET /batches/{batch_id}. "
        "Строки с ошибками — в GET /batches/{batch_id}/errors.xlsx со столбцом «Комментарий». "
        "Исправленный файл можно загрузить повторно: уже созданные договоры не задублируются."
    ),
    responses={422: {"description": "Файл не принят целиком (нет листов, ошибки на листе «Объект»)"}},
)
async def upload_bulk_excel(
    file: UploadFile = File(..., description="Excel по шаблону GET /contracts/subleases/template"),
    user_id: str | None = Query(None, title="Кто создал"),
    service: LeaseBatchService = Depends(get_batch_service),
) -> BatchAcceptedResponse:
    if not (file.filename or "").endswith((".xlsx", ".xlsm")):
        raise HTTPException(status_code=status.HTTP_400_BAD_REQUEST, detail="Нужен файл .xlsx")
    try:
        summary = await service.accept(await file.read(), file.filename, user_id)
    except BatchFileRejected as e:
        raise HTTPException(status_code=status.HTTP_422_UNPROCESSABLE_ENTITY, detail=e.errors)
    return BatchAcceptedResponse(
        **summary.__dict__,
        status_url=_status_url(summary.batch_id),
        errors_url=_errors_url(summary.batch_id) if summary.rejected else None,
    )


@router.get("/batches/{batch_id}", response_model=BatchStatusResponse, summary="Статус пакетной загрузки")
async def get_batch_status(
    batch_id: uuid.UUID,
    service: LeaseBatchService = Depends(get_batch_service),
) -> BatchStatusResponse:
    result = await service.status(batch_id)
    if result is None:
        raise HTTPException(status_code=status.HTTP_404_NOT_FOUND, detail="Пачка не найдена")
    return BatchStatusResponse(
        **{**result.__dict__, "errors": [BatchRowErrorResponse(**e.__dict__) for e in result.errors]},
        errors_url=_errors_url(batch_id) if result.errors else None,
    )


@router.get(
    "/batches/{batch_id}/errors.xlsx",
    summary="Excel со строками, которые не создались, и столбцом «Комментарий»",
    response_class=Response,
    responses={200: {"content": {_XLSX_TYPE: {}}}},
)
async def download_batch_errors(
    batch_id: uuid.UUID,
    service: LeaseBatchService = Depends(get_batch_service),
) -> Response:
    result = await service.errors_workbook(batch_id)
    if result is None:
        raise HTTPException(status_code=status.HTTP_404_NOT_FOUND, detail="Пачка не найдена")
    content, filename = result
    return Response(
        content=content,
        media_type=_XLSX_TYPE,
        headers={"Content-Disposition": f"attachment; filename=\"errors.xlsx\"; filename*=UTF-8''{quote(filename)}"},
    )


@router.post(
    "/batches/{batch_id}/retry",
    summary="[Администратор] Повторить строки, упавшие при создании (status=error)",
    description=(
        "Для поддержки после аварии (MinIO/Gotenberg/БД были недоступны). "
        "Строки rejected не повторяются — их нужно исправить в errors.xlsx и загрузить заново."
    ),
)
async def retry_batch(
    batch_id: uuid.UUID,
    service: LeaseBatchService = Depends(get_batch_service),
) -> dict:
    return {"requeued": await service.retry(batch_id)}
