"""
LeaseContractService — создание договора аренды/субаренды, готового к подписанию.

Результат одного create():
  MinIO:  {PAYDA}/<имя>.docx (источник), {PAYDA}/<имя>.pdf, {PAYDA}/<имя> <uuid>.xml
  БД:     lease_contracts, documents (pdf + xml, статус SENT), participants (SELLER + BUYER)

Дальше договор подписывается существующим механизмом payda_edo → consumer-parks (upload_signs):
SELLER (арендодатель) подписывает первым, затем BUYER (арендатор).
"""

from __future__ import annotations

import asyncio
import hashlib
import logging
import time
import uuid
from dataclasses import dataclass
from datetime import date
from decimal import Decimal
from io import BytesIO

from sqlalchemy import select
from sqlalchemy.exc import IntegrityError
from sqlalchemy.ext.asyncio import AsyncSession, async_sessionmaker
from typing import Callable

from clients import GotenbergClient
from db.minio import MinioStorage
from db.models import Document, LeaseContract, Participant
from repo import CONTRACT_TEMPLATES, ContractRenderer, renderer_for
from schemas import ORG_TYPES_WITH_BIN, SubleaseContractData
from schemas.service_agreement import ContractDate
from services.lease_xml import build_lease_xml

logger = logging.getLogger(__name__)

CATEGORY_SUBLEASE = "SUBLEASE"
CATEGORY_LEASE = "LEASE"
SOURCE = "INTERNAL_EDO"
STATUS_SENT = "SENT"

_DOCX_TYPE = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
_TITLES = {CATEGORY_SUBLEASE: "Договор субаренды", CATEGORY_LEASE: "Договор аренды"}


class LeaseContractConflict(Exception):
    """Номер договора у этого арендодателя уже занят другим арендатором."""


@dataclass
class LeaseContractResult:
    id: int
    category: str
    contract_number: str
    status: str | None
    pdf_id: uuid.UUID
    xml_id: uuid.UUID
    pdf_url: str
    created: bool  # False — договор уже существовал (повторный запрос)


def _to_date(d: ContractDate) -> date:
    return date(int(d.year), int(d.month_num), int(d.day))


def _to_decimal(value: str) -> Decimal:
    return Decimal(value.replace(" ", "").replace("\xa0", ""))


def _party_xml(party, tin: str, name: str, phone: str | None = None) -> dict:
    """Сторона в XML. Блок director — директор по тексту договора (только ТОО/АО).
    Кто фактически подписал, в XML не пишем: это видно из сертификата (signs.signer_data);
    consumer сверяет с участником только ИИН/БИН из сертификата."""
    fields = {"org_type": party.org_type, "tin": tin, "name": name}
    if phone is not None:
        fields["phone"] = phone
    fields["address"] = party.address_ru
    if party.org_type in ORG_TYPES_WITH_BIN:
        director = {"name": party.director_kz}
        if party.director_iin:
            director["iin"] = party.director_iin
        fields["director"] = director
    return fields


def lease_xml_fields(lease: LeaseContract, data: SubleaseContractData, pdf_doc_id, payda, pdf_hash: str) -> dict:
    """Содержимое XML, которое подписывают стороны."""
    return {
        "id": lease.id,
        "category": lease.category,
        "document_number": lease.contract_number,
        "document_date": lease.signed_date.isoformat(),
        "seller": _party_xml(data.lessor, lease.seller_tin, lease.seller_name),
        "buyer": _party_xml(data.tenant, lease.buyer_tin, lease.buyer_name, phone=lease.buyer_phone),
        # Основной договор аренды — только для субаренды
        **({"head_lease": {
            "number": data.building.lease_contract_num,
            "date": data.building.lease_contract_date,
            "owner_name": data.building.owner_name_ru,
        }} if lease.category == CATEGORY_SUBLEASE else {}),
        "period": {
            "start_date": lease.start_date.isoformat(),
            "end_date": lease.end_date.isoformat(),
        },
        "premises": {
            "building_name": data.building.building_name_ru,
            "address": data.building.building_address_ru,
            "floor": lease.floor,
            "room_type": lease.room_type,
            "area": lease.area,
        },
        "payment": {
            "currency": "KZT",
            "rent_per_sqm": lease.rent_per_sqm,
            "rent_total": lease.rent_total,
            "communal": lease.communal,
        },
        "pdf_id": pdf_doc_id,
        "payda_id": payda,
        "hash": pdf_hash,
    }


class LeaseContractService:
    def __init__(
        self,
        storage: MinioStorage,
        converter: GotenbergClient,
        session_factory: async_sessionmaker[AsyncSession],
        renderer_provider: Callable[[str], ContractRenderer] = renderer_for,
    ):
        self._renderer_provider = renderer_provider  # шаблон .docx по типу договора
        self._storage = storage
        self._converter = converter
        self._session_factory = session_factory

    async def create(self, data: SubleaseContractData, created_by: str | None = None) -> LeaseContractResult:
        category = data.category
        seller_tin, buyer_tin = data.lessor.iin, data.tenant.iin

        # Идемпотентность: повторный запрос возвращает уже созданный договор
        async with self._session_factory() as session:
            existing = await self._find(session, category, seller_tin, data.contract_number)
            if existing is not None:
                return await self._existing_result(session, existing, buyer_tin)

        start = time.time()
        payda = uuid.uuid4()
        signed = _to_date(data.signed_date)
        title = f"{_TITLES[category]} № {data.contract_number} от {signed:%d.%m.%Y} {buyer_tin}".replace("/", "_")

        docx = self._renderer_provider(category).render(data).getvalue()
        pdf = await self._converter.docx_to_pdf(docx, f"{title}.docx")
        pdf_hash = hashlib.sha256(pdf).hexdigest()

        pdf_name = f"{title}.pdf"
        xml_name = f"{title} {uuid.uuid4()}.xml"
        docx_key = f"{payda}/{title}.docx"
        pdf_key = f"{payda}/{pdf_name}"
        xml_key = f"{payda}/{xml_name}"

        uploaded: list[str] = []
        try:
            await self._upload(docx_key, docx, _DOCX_TYPE, uploaded)
            await self._upload(pdf_key, pdf, "application/pdf", uploaded)

            async with self._session_factory() as session, session.begin():
                lease = self._build_lease(data, category, docx_key)
                session.add(lease)
                await session.flush()  # нужен lease.id для XML

                pdf_doc_id = uuid.uuid4()
                post_data = self._post_data(lease, data)
                post_data["document_id"] = str(pdf_doc_id)

                xml = build_lease_xml(lease_xml_fields(lease, data, pdf_doc_id, payda, pdf_hash))
                await self._upload(xml_key, xml, "application/xml", uploaded)

                pdf_doc = self._document(
                    pdf_doc_id, payda, category, "pdf", pdf_key, pdf_name, post_data, created_by, pdf_hash,
                )
                xml_doc = self._document(
                    uuid.uuid4(), None, category, "xml", xml_key, xml_name, post_data, created_by, None,
                )
                pdf_doc.process_time = str(time.time() - start)
                session.add_all([pdf_doc, xml_doc])
                await session.flush()

                # Участники привязаны к PDF-документу: по нему payda_edo ищет подписанта и меняет статус
                session.add_all([
                    Participant(document_id=pdf_doc.id, tin=seller_tin, type="SELLER"),
                    Participant(document_id=pdf_doc.id, tin=buyer_tin, type="BUYER"),
                ])
                lease.pdf_id, lease.xml_id = pdf_doc.id, xml_doc.id

        except IntegrityError:
            # Параллельный запрос успел создать договор с тем же номером
            await self._cleanup(uploaded)
            async with self._session_factory() as session:
                existing = await self._find(session, category, seller_tin, data.contract_number)
                if existing is None:
                    raise
                return await self._existing_result(session, existing, buyer_tin)
        except Exception:
            await self._cleanup(uploaded)
            raise

        logger.info("Договор %s №%s создан: lease_id=%s pdf_id=%s", category, lease.contract_number, lease.id, lease.pdf_id)
        return LeaseContractResult(
            id=lease.id,
            category=category,
            contract_number=lease.contract_number,
            status=STATUS_SENT,
            pdf_id=lease.pdf_id,
            xml_id=lease.xml_id,
            pdf_url=self._storage.file_url(pdf_key),
            created=True,
        )

    # ── helpers ───────────────────────────────────────────────────────────────

    @staticmethod
    async def _find(session: AsyncSession, category: str, seller_tin: str, number: str) -> LeaseContract | None:
        statement = select(LeaseContract).where(
            LeaseContract.category == category,
            LeaseContract.seller_tin == seller_tin,
            LeaseContract.contract_number == number,
        )
        return (await session.execute(statement)).scalars().first()

    async def _existing_result(self, session: AsyncSession, lease: LeaseContract, buyer_tin: str) -> LeaseContractResult:
        if lease.buyer_tin != buyer_tin:
            raise LeaseContractConflict(
                f"Договор №{lease.contract_number} уже создан для другого арендатора (ИИН/БИН {lease.buyer_tin})"
            )
        pdf_doc = await session.get(Document, lease.pdf_id)
        object_name = f"{pdf_doc.PAYDA}/{pdf_doc.filename}"
        return LeaseContractResult(
            id=lease.id,
            category=lease.category,
            contract_number=lease.contract_number,
            status=pdf_doc.current_status,
            pdf_id=lease.pdf_id,
            xml_id=lease.xml_id,
            pdf_url=self._storage.file_url(object_name),
            created=False,
        )

    @staticmethod
    def _build_lease(data: SubleaseContractData, category: str, docx_key: str) -> LeaseContract:
        f = data.financial
        return LeaseContract(
            category=category,
            contract_number=data.contract_number,
            signed_date=_to_date(data.signed_date),
            start_date=_to_date(data.start_date),
            end_date=_to_date(data.end_date),
            seller_tin=data.lessor.iin,
            seller_name=data.lessor.name_ru,
            buyer_tin=data.tenant.iin,
            buyer_name=data.tenant.name_ru,
            buyer_phone=data.tenant.phone or None,
            floor=f.floor,
            room_type=f.room_type,
            area=_to_decimal(f.area),
            rent_per_sqm=_to_decimal(f.rent_per_sqm),
            rent_total=_to_decimal(f.rent_total),
            communal=_to_decimal(f.communal),
            data=data.model_dump(mode="json"),
            docx_path=docx_key,
            template_code=CONTRACT_TEMPLATES[category].code,
            template_version=CONTRACT_TEMPLATES[category].version,
        )

    @staticmethod
    def _post_data(lease: LeaseContract, data: SubleaseContractData) -> dict:
        """post_data строк documents. Ключи document_number/document_date и seller/buyer — как у других документов."""
        return {
            "lease_contract_id": lease.id,
            "document_number": lease.contract_number,
            "document_date": lease.signed_date.isoformat(),
            "seller": {"iin": lease.seller_tin, "name": lease.seller_name, "address": data.lessor.address_ru},
            "buyer": {"iin": lease.buyer_tin, "name": lease.buyer_name, "phone": lease.buyer_phone},
            "contract": {"number": lease.contract_number, "date": lease.signed_date.isoformat()},
            "period": {"start": lease.start_date.isoformat(), "end": lease.end_date.isoformat()},
            "amount_total": str(lease.rent_total),
        }

    def _document(
        self,
        doc_id: uuid.UUID,
        payda: uuid.UUID | None,
        category: str,
        extension: str,
        object_name: str,
        filename: str,
        post_data: dict,
        created_by: str | None,
        hash_sha256: str | None,
    ) -> Document:
        doc = Document(
            id=doc_id,
            post_data=post_data,
            filepath=f"{self._storage.base_url}/{self._storage.bucket}/{object_name}",
            category=category,
            extension_type=extension,
            created_by=created_by,
            current_status=STATUS_SENT,
            filename=filename,
            source=SOURCE,
            hash_sha256=hash_sha256,
        )
        if payda is not None:
            doc.PAYDA = payda
        return doc

    async def _upload(self, object_name: str, content: bytes, content_type: str, uploaded: list[str]) -> None:
        await asyncio.to_thread(self._storage.upload, object_name, BytesIO(content), content_type)
        uploaded.append(object_name)

    async def _cleanup(self, object_names: list[str]) -> None:
        for name in object_names:
            try:
                await asyncio.to_thread(self._storage.delete, name)
            except Exception:
                logger.exception("Не удалось удалить %s из MinIO после ошибки", name)
