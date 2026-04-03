from pydantic_settings import BaseSettings, SettingsConfigDict


class MinioSettings(BaseSettings):
    model_config = SettingsConfigDict(
        env_prefix='minio_',
        env_file='.env',
        extra='ignore',
    )

    endpoint: str
    public_endpoint: str | None = None
    access_key: str
    secret_key: str
    bucket: str
    secure: bool = False
