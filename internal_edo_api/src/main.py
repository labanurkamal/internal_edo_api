"""
Точка входа FastAPI.
"""

from fastapi import FastAPI
from api import router_v1 as contracts_router

app = FastAPI(
    title="Contract Generator",
    docs_url="/api/docs",
    redoc_url="/api/redoc",
    description="Генерация договоров субаренды и аренды",
    version="1.0.0",
)

app.include_router(contracts_router)


@app.get("/health")
async def health():
    return {"status": "ok"}