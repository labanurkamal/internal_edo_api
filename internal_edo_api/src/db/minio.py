"""
Работа с MinIO: загрузка файлов и получение URL.
Единственная ответственность — хранение. Бизнес-логика снаружи.
"""

from __future__ import annotations
from io import BytesIO
from datetime import timedelta
import json

from minio import Minio
from minio.error import S3Error

from core.config import Settings


class MinioStorage:
    """Обёртка над MinIO. Зависимость инжектируется — легко мокать в тестах."""

    def __init__(self, settings: Settings):
        self._client = Minio(
            endpoint=settings.minio.endpoint,
            access_key=settings.minio.access_key,
            secret_key=settings.minio.secret_key,
            secure=settings.minio.secure,
        )
        self._bucket       = settings.minio.bucket
        self._public_read  = settings.minio.public_read
        self._base_url     = settings.minio.base_url.rstrip("/")
        self._presigned_ttl = timedelta(hours=settings.minio.presigned_ttl_hours)
        self._ensure_bucket()

    def _ensure_bucket(self) -> None:
        """
        Создаёт bucket если не существует.
        Если MINIO_PUBLIC_READ=true — ставит публичный read policy:
          - ссылки без подписи и срока действия
          - открываются в Google Docs Viewer / браузере напрямую
        """
        if not self._client.bucket_exists(self._bucket):
            self._client.make_bucket(self._bucket)

        if self._public_read:
            policy = {
                "Version": "2012-10-17",
                "Statement": [{
                    "Effect": "Allow",
                    "Principal": {"AWS": ["*"]},
                    "Action": ["s3:GetObject"],
                    "Resource": [f"arn:aws:s3:::{self._bucket}/*"],
                }],
            }
            self._client.set_bucket_policy(self._bucket, json.dumps(policy))

    def upload(self, object_name: str, data: BytesIO, content_type: str = "application/vnd.openxmlformats-officedocument.wordprocessingml.document") -> str:
        """Загружает файл в MinIO, возвращает object_name."""
        length = data.getbuffer().nbytes
        data.seek(0)
        self._client.put_object(
            bucket_name=self._bucket,
            object_name=object_name,
            data=data,
            length=length,
            content_type=content_type,
        )
        return object_name

    def file_url(self, object_name: str) -> str:
        """
        Возвращает URL файла:
          - public_read=True  → постоянная публичная ссылка (без подписи)
                                 https://storage.edo.kz/contracts/subleases/file.pdf
          - public_read=False → presigned URL (истекает через presigned_ttl_hours)
        """
        if self._public_read:
            return f"{self._base_url}/{self._bucket}/{object_name}"
        return self._client.presigned_get_object(
            bucket_name=self._bucket,
            object_name=object_name,
            expires=self._presigned_ttl,
        )

    # Оставляем для обратной совместимости
    def presigned_url(self, object_name: str) -> str:
        return self.file_url(object_name)

    def delete(self, object_name: str) -> None:
        self._client.remove_object(self._bucket, object_name)