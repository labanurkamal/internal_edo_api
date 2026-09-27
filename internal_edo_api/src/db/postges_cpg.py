import asyncpg
import json

from core.config import settings


class PostgresDB:
    def __init__(self, dsn: str, pool_size_min: int = 5, pool_size_max: int = 10):
        self.dsn = dsn
        self.pool_size_min = pool_size_min
        self.pool_size_max = pool_size_max
        self._pool: asyncpg.Pool | None = None

    @staticmethod
    async def _init_connection(conn: asyncpg.Connection) -> None:
        for typename in ("json", "jsonb"):
            await conn.set_type_codec(
                typename,
                encoder=json.dumps,
                decoder=json.loads,
                schema="pg_catalog",
                format="text",
            )

    async def connect(self) -> None:
        self._pool = await asyncpg.create_pool(  # type: ignore[misc]
            dsn=self.dsn,
            min_size=self.pool_size_min,
            max_size=self.pool_size_max,
            timeout=30,
            command_timeout=30,
            max_inactive_connection_lifetime=300,
            init=self._init_connection,
        )

    async def get_pool(self) -> asyncpg.Pool:
        if self._pool is None:
            await self.connect()

        return self._pool  # type: ignore

    async def close(self) -> None:
        if self._pool is not None:
            await self._pool.close()
            self._pool = None


db_payda = PostgresDB(
    dsn=settings.db.dsn, pool_size_min=5, pool_size_max=10
)
