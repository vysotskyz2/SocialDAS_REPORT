from io import BytesIO
from datetime import timedelta

from miniopy_async import Minio
from loguru import logger


class MinioClient:
    def __init__(
        self,
        endpoint: str,
        access_key: str,
        secret_key: str,
        bucket: str,
        secure: bool = False,
    ) -> None:
        self._client = Minio(
            endpoint,
            access_key=access_key,
            secret_key=secret_key,
            secure=secure,
        )
        self._bucket = bucket

    async def ensure_bucket(self) -> None:
        if not await self._client.bucket_exists(self._bucket):
            await self._client.make_bucket(self._bucket)
            logger.info(f"Created MinIO bucket: {self._bucket}")

    async def upload_file(
        self,
        key: str,
        data: bytes,
        content_type: str = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    ) -> None:
        await self._client.put_object(
            self._bucket,
            key,
            BytesIO(data),
            length=len(data),
            content_type=content_type,
        )

    async def get_download_url(self, key: str, expiry: int = 3600) -> str:
        return await self._client.presigned_get_object(
            self._bucket,
            key,
            expires=timedelta(seconds=expiry),
        )

    async def delete_file(self, key: str) -> None:
        await self._client.remove_object(self._bucket, key)
