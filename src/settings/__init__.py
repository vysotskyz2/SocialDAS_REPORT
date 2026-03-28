from src.settings.db import DatabaseConfig
from src.settings.minio import MinioSettings
from src.settings.app import AppSettings


class Settings:
    db: DatabaseConfig = DatabaseConfig()
    minio: MinioSettings = MinioSettings()
    app: AppSettings = AppSettings()


settings = Settings()
