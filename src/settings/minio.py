from pydantic_settings import BaseSettings, SettingsConfigDict


class MinioSettings(BaseSettings):
    model_config = SettingsConfigDict(
        env_prefix='minio_',
        env_file='.env',
        extra='ignore',
    )

    endpoint: str
    access_key: str
    secret_key: str
    bucket: str
    secure: bool = False
