"""
Точка входа FastAPI.
"""

from contextlib import asynccontextmanager

from fastapi import FastAPI
from api import router_v1 as contracts_router
from api.v1.lease import get_batch_service, router as lease_router
from core.config import settings
from services.lease_batch import LeaseBatchWorker


@asynccontextmanager
async def lifespan(app: FastAPI):
    # Воркер пакетной загрузки договоров: очередь в Postgres (lease_batch_items)
    worker = LeaseBatchWorker(get_batch_service()) if settings.lease_worker_enabled else None
    if worker:
        worker.start()
    yield
    if worker:
        await worker.stop()


app = FastAPI(
    title="Contract Generator",
    docs_url="/api/docs",
    redoc_url="/api/redoc",
    description="Генерация договоров субаренды и аренды",
    version="1.0.0",
    lifespan=lifespan,
)

app.include_router(contracts_router)
app.include_router(lease_router)


@app.get("/health")
async def health():
    return {"status": "ok"}
