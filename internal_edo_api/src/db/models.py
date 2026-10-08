"""
ORM-модели.

Document и Participant — общие таблицы EDO, которыми пользуются payda_edo и doc-creator-edo-consumer.
Их схема задана в БД, здесь она только повторена: колонки не добавлять и не переименовывать.
У documents.id / created / current_status в БД нет значений по умолчанию — их заполняет код.

LeaseContract, LeaseBatch, LeaseBatchItem — таблицы internal_edo_api (deploy/sql/001_lease_contracts.sql).
"""

from __future__ import annotations

import uuid
from datetime import date, datetime
from decimal import Decimal

from sqlalchemy import Boolean, Date, DateTime, ForeignKey, Integer, Numeric, String, Text, UniqueConstraint
from sqlalchemy.dialects.postgresql import JSONB, UUID
from sqlalchemy.orm import DeclarativeBase, Mapped, mapped_column


class Base(DeclarativeBase):
    pass


# ── Общие таблицы EDO ─────────────────────────────────────────────────────────

class Document(Base):
    __tablename__ = "documents"

    id: Mapped[uuid.UUID] = mapped_column(UUID(as_uuid=True), primary_key=True, default=uuid.uuid4)
    post_data: Mapped[dict | None] = mapped_column(JSONB)
    created: Mapped[datetime | None] = mapped_column(DateTime, default=datetime.now)
    updated: Mapped[datetime | None] = mapped_column(DateTime, default=datetime.now, onupdate=datetime.now)
    isAccessibleByContractor: Mapped[bool | None] = mapped_column(Boolean, default=False)
    isPublicToInternet: Mapped[bool | None] = mapped_column(Boolean, default=False)
    filepath: Mapped[str | None] = mapped_column(String)
    category: Mapped[str | None] = mapped_column(String)
    extension_type: Mapped[str | None] = mapped_column(String)
    created_by: Mapped[str | None] = mapped_column(String)
    current_status: Mapped[str | None] = mapped_column(String)
    process_time: Mapped[str | None] = mapped_column(String)
    filename: Mapped[str | None] = mapped_column(String)
    PAYDA: Mapped[uuid.UUID | None] = mapped_column(UUID(as_uuid=True), default=uuid.uuid4)
    source: Mapped[str | None] = mapped_column(String)
    hash_sha256: Mapped[str | None] = mapped_column(String)
    error_log: Mapped[str | None] = mapped_column(String)


class Participant(Base):
    __tablename__ = "participants"

    id: Mapped[int] = mapped_column(Integer, primary_key=True)
    document_id: Mapped[uuid.UUID | None] = mapped_column(ForeignKey("documents.id", ondelete="CASCADE"))
    tin: Mapped[str | None] = mapped_column(String)
    type: Mapped[str | None] = mapped_column(String)  # SELLER / BUYER


# ── Таблицы internal_edo_api ──────────────────────────────────────────────────

class LeaseContract(Base):
    """Договор аренды/субаренды. Аналог awps у акта: данные договора + ссылки на PDF и XML."""

    __tablename__ = "lease_contracts"
    __table_args__ = (
        # Номер договора уникален в пределах арендодателя — это ключ идемпотентности
        UniqueConstraint("category", "seller_tin", "contract_number", name="uq_lease_contracts_number"),
    )

    id: Mapped[int] = mapped_column(Integer, primary_key=True)
    category: Mapped[str] = mapped_column(String(16))  # LEASE / SUBLEASE
    contract_number: Mapped[str] = mapped_column(String)
    signed_date: Mapped[date] = mapped_column(Date)
    start_date: Mapped[date] = mapped_column(Date)
    end_date: Mapped[date] = mapped_column(Date)

    seller_tin: Mapped[str] = mapped_column(String(12))
    seller_name: Mapped[str] = mapped_column(String)
    buyer_tin: Mapped[str] = mapped_column(String(12))
    buyer_name: Mapped[str] = mapped_column(String)
    buyer_phone: Mapped[str | None] = mapped_column(String)

    floor: Mapped[str | None] = mapped_column(String)
    room_type: Mapped[str | None] = mapped_column(String)
    area: Mapped[Decimal] = mapped_column(Numeric(12, 2))
    rent_per_sqm: Mapped[Decimal] = mapped_column(Numeric(14, 2))
    rent_total: Mapped[Decimal] = mapped_column(Numeric(14, 2))
    communal: Mapped[Decimal] = mapped_column(Numeric(14, 2))

    data: Mapped[dict] = mapped_column(JSONB)  # полные данные для шаблона (SubleaseContractData)
    docx_path: Mapped[str | None] = mapped_column(String)  # исходный .docx в MinIO, не подписывается
    # Каким шаблоном .docx создан договор (repo.CONTRACT_TEMPLATES): нужен, когда юристы меняют текст
    template_code: Mapped[str | None] = mapped_column(String(16))
    template_version: Mapped[int | None] = mapped_column(Integer)
    pdf_id: Mapped[uuid.UUID | None] = mapped_column(ForeignKey("documents.id"))
    xml_id: Mapped[uuid.UUID | None] = mapped_column(ForeignKey("documents.id"))
    created: Mapped[datetime] = mapped_column(DateTime, default=datetime.now)


class LeaseBatch(Base):
    """Пакетная загрузка из Excel."""

    __tablename__ = "lease_batches"

    id: Mapped[uuid.UUID] = mapped_column(UUID(as_uuid=True), primary_key=True, default=uuid.uuid4)
    category: Mapped[str] = mapped_column(String(16))
    filename: Mapped[str | None] = mapped_column(String)
    source_path: Mapped[str | None] = mapped_column(String)  # загруженный Excel в MinIO
    total: Mapped[int] = mapped_column(Integer, default=0)
    created_by: Mapped[str | None] = mapped_column(String)
    created: Mapped[datetime] = mapped_column(DateTime, default=datetime.now)


class LeaseBatchItem(Base):
    """Одна строка Excel. Очередь фоновой обработки: воркер берёт status='pending'."""

    __tablename__ = "lease_batch_items"

    id: Mapped[int] = mapped_column(Integer, primary_key=True)
    batch_id: Mapped[uuid.UUID] = mapped_column(ForeignKey("lease_batches.id", ondelete="CASCADE"))
    row_number: Mapped[int] = mapped_column(Integer)  # номер строки в исходном Excel
    payload: Mapped[dict | None] = mapped_column(JSONB)  # SubleaseContractData, если строка прошла проверку
    status: Mapped[str] = mapped_column(String(16))  # rejected / pending / processing / done / error
    error: Mapped[str | None] = mapped_column(Text)
    attempts: Mapped[int] = mapped_column(Integer, default=0)
    lease_contract_id: Mapped[int | None] = mapped_column(ForeignKey("lease_contracts.id"))
    created: Mapped[datetime] = mapped_column(DateTime, default=datetime.now)
    updated: Mapped[datetime] = mapped_column(DateTime, default=datetime.now, onupdate=datetime.now)
