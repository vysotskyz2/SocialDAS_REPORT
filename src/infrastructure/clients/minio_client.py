from io import BytesIO
from datetime import timedelta
from miniopy_async import Minio
from miniopy_async.error import S3Error
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
        try:
            await self._client.make_bucket(self._bucket)
            logger.info(f"Создан MinIO bucket: {self._bucket}")
        except S3Error as e:
            if e.code != "BucketAlreadyOwnedByYou":
                raise

    async def upload_file(
        self,
        key: str,
        data: bytes,
        content_type: str = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    ) -> None:
        try:
            await self._client.put_object(
                self._bucket, key, BytesIO(data),
                length=len(data), content_type=content_type,
            )
        except S3Error as e:
            logger.error(f"Ошибка при загрузке {key}: {e}")
            raise

    async def get_download_url(self, key: str, expiry: int = 3600) -> str:
        return await self._client.presigned_get_object(
            self._bucket,
            key,
            expires=timedelta(seconds=expiry),
        )

    async def delete_file(self, key: str) -> None:
        try:
            await self._client.remove_object(self._bucket, key)
        except S3Error as e:
            logger.error(f"Ошибка при удалении {key}: {e}")
            raise
