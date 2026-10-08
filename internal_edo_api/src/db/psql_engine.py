import asyncio
import uuid

from sqlalchemy import NullPool
from sqlalchemy.ext.asyncio import async_sessionmaker, AsyncSession, create_async_engine

from core.config import settings

def _async_dsn(dsn: str) -> str:
    """DB_URL общий с asyncpg-пулом (postgresql://...); SQLAlchemy нужен явный драйвер."""
    for prefix in ("postgresql://", "postgres://"):
        if dsn.startswith(prefix):
            return "postgresql+asyncpg://" + dsn[len(prefix):]
    return dsn


engine = create_async_engine(
    url=_async_dsn(settings.db.dsn),
    pool_pre_ping=True,
    echo=False,
    pool_recycle=1800,
    poolclass=NullPool,
    connect_args={
        "statement_cache_size": 0,
        "command_timeout": 60,
        "prepared_statement_name_func": lambda: f"__asyncpg_{uuid.uuid4().hex}__",
    },
)

session = async_sessionmaker(
    bind=engine,
    expire_on_commit=False,
    class_=AsyncSession
)

