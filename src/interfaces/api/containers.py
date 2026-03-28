from contextvars import ContextVar
from typing import AsyncGenerator

from dependency_injector import containers, providers
from httpx import AsyncClient, AsyncHTTPTransport, Limits
from sqlalchemy.ext.asyncio import AsyncSession

from src.infrastructure.clients.analytics_client import AnalyticsClient
from src.infrastructure.clients.minio_client import MinioClient
from src.infrastructure.repositories.report_repository import ReportRepository
from src.application.services.report_service import ReportService
from src.application.services.excel_builder import ExcelBuilder
from src.settings import settings

db_session_context: ContextVar[AsyncSession] = ContextVar("db_session_context")


def get_db_session() -> AsyncSession:
    return db_session_context.get()


async def init_async_client() -> AsyncGenerator[AsyncClient, None]:
    transport = AsyncHTTPTransport(retries=3, verify=True, http2=True)
    limits = Limits(
        max_keepalive_connections=5,
        max_connections=20,
        keepalive_expiry=5.0,
    )
    async with AsyncClient(
        timeout=30.0,
        transport=transport,
        limits=limits,
    ) as client:
        yield client


class Container(containers.DeclarativeContainer):
    wiring_config = containers.WiringConfiguration(
        modules=[
            "src.interfaces.api.routers.reports",
            "src.interfaces.api.app",
        ]
    )

    config = providers.Object(settings)

    db_session = providers.Callable(get_db_session)

    http_client = providers.Resource(init_async_client)

    analytics_client = providers.Factory(
        AnalyticsClient,
        http_client=http_client,
        base_url=settings.app.analytics_service_url,
    )

    minio_client = providers.Singleton(
        MinioClient,
        endpoint=settings.minio.endpoint,
        access_key=settings.minio.access_key,
        secret_key=settings.minio.secret_key,
        bucket=settings.minio.bucket,
        secure=settings.minio.secure,
    )

    excel_builder = providers.Singleton(ExcelBuilder)

    report_repository = providers.Factory(
        ReportRepository,
        session=db_session,
    )

    report_service = providers.Factory(
        ReportService,
        repository=report_repository,
        analytics_client=analytics_client,
        minio_client=minio_client,
        excel_builder=excel_builder,
        download_url_expiry=settings.app.download_url_expiry,
    )
