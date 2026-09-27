import os
from logging import config as logging_config

from dotenv import find_dotenv
from pydantic import Field
from pydantic_settings import BaseSettings, SettingsConfigDict

from core.logger import LOGGING

logging_config.dictConfig(LOGGING)

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))


class ModelConfigs(BaseSettings):
    model_config = SettingsConfigDict(
        env_file=find_dotenv(filename=".env"),
        env_file_encoding="utf-8",
        extra="ignore"
    )


class DBSettings(ModelConfigs):
    dsn: str = Field(..., alias="DB_URL")


class MinioSettings(ModelConfigs):
    access_key: str = Field(..., alias="MINIO_ACCESS_KEY")
    secret_key: str = Field(..., alias="MINIO_SECRET_KEY")
    endpoint: str = Field(..., alias="MINIO_ENDPOINT")
    url: str = Field(..., alias="MINIO_URL")
    bucket: str = Field(..., alias="MINIO_BUCKET")
    presigned_ttl_hours: int = 24
    secure: bool = False
    public_read: bool = True
    base_url: str = Field(..., alias="MINIO_DOMAIN")

class Settings(BaseSettings):
    db: DBSettings = DBSettings()
    minio: MinioSettings = MinioSettings()


settings = Settings()
