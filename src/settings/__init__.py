from src.settings.db import DatabaseConfig
from src.settings.minio import MinioSettings
from src.settings.app import AppSettings
from src.settings.redis import RedisSettings
from pydantic_settings import BaseSettings


class Settings(BaseSettings):
    db: DatabaseConfig = DatabaseConfig()
    minio: MinioSettings = MinioSettings()
    app: AppSettings = AppSettings()
    redis: RedisSettings = RedisSettings()


settings = Settings()
