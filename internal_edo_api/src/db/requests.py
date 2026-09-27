from enum import Enum
from typing import Mapping, Any

import httpx


class Methods(Enum):
    GET = "GET"
    POST = "POST"
    PUT = "PUT"
    PATCH = "PATCH"
    DELETE = "DELETE"
    HEAD = "HEAD"
    OPTIONS = "OPTIONS"


class AsyncClient:
    def __init__(
        self,
        base_url: str,
        *,
        headers: dict[str, str] | None = None,
        timeout: httpx.Timeout | None = None,
        limits: httpx.Limits | None = None,
        http2: bool = True,
    ):
        default_headers = {
            "Accept": "application/json",
        }
        self.client = httpx.AsyncClient(
            base_url=base_url,
            headers={
                **default_headers,
                **(dict(headers) if headers else {}),
            },
            timeout=timeout
            or httpx.Timeout(
                connect=5.0,
                read=20.0,
                write=20.0,
                pool=5.0,
            ),
            limits=limits
            or httpx.Limits(
                max_connections=100,
                max_keepalive_connections=20,
                keepalive_expiry=30.0,
            ),
            http2=http2,
        )

    async def __aexit__(self, *_):
        await self.client.aclose()

    async def close(self):
        await self.client.aclose()

    async def send(
        self,
        method: Methods,
        endpoint: str,
        *,
        headers: Mapping[str, str] | None = None,
        json: Any | None = None,
        data: Any | None = None,
    ):
        response = await self.client.request(
            method=method.value,
            url=endpoint,
            headers=headers,
            json=json,
            data=data,
        )
        response.raise_for_status()
        return response
