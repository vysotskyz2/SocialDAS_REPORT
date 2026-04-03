import asyncio
import uuid
from io import BytesIO
from loguru import logger
from src.infrastructure.clients.analytics_client import AnalyticsClient
from src.infrastructure.clients.minio_client import MinioClient
from src.infrastructure.repositories.report_repository import ReportRepository
from src.infrastructure.schemas.report import (
    ReportRequest,
    ReportResponse,
    ReportType,
    ReportListResponse,
)
from src.application.services.excel_builder import ExcelBuilder


class ReportService:
    def __init__(
        self,
        repository: ReportRepository,
        analytics_client: AnalyticsClient,
        minio_client: MinioClient,
        excel_builder: ExcelBuilder,
        download_url_expiry: int = 3600,
    ) -> None:
        self._repo = repository
        self._analytics = analytics_client
        self._minio = minio_client
        self._excel = excel_builder
        self._download_url_expiry = download_url_expiry

    async def generate_report(
        self, user_id: str, request: ReportRequest
    ) -> ReportResponse:
        from src.infrastructure.tasks.reports import generate_report_task

        report = await self._repo.create(
            user_id=user_id,
            platform=request.platform.value,
            account_id=request.account_id,
            report_type=request.report_type.value,
        )

        await generate_report_task.kiq(
            report_id=report.id,
            user_id=user_id,
            request_data=request.model_dump(),
        )

        return self._to_response(report)

    async def process_report_generation(
        self, report_id: uuid.UUID, user_id: str, request: ReportRequest
    ) -> None:
        report = await self._repo.get_by_id(report_id)
        if not report:
            logger.error(f"Отчет {report_id} не найден")
            return

        try:
            data = await self._fetch_data(request)

            excel_bytes = self._excel.build(
                platform=request.platform.value,
                account_id=request.account_id,
                report_type=request.report_type.value,
                data=data,
            )

            timestamp = report.created_at.strftime("%Y%m%d_%H%M%S")
            file_name = (
                f"{request.platform.value}_{request.account_id}"
                f"_{request.report_type.value}_{timestamp}.xlsx"
            )
            file_key = f"{user_id}/{report.id}/{file_name}"

            await self._minio.upload_file(file_key, excel_bytes)
            await self._repo.update_completed(report.id, file_key, file_name)

        except Exception as e:
            logger.error(f"Ошибка генерации отчета {report.id}: {e}")
            await self._repo.update_failed(report.id, str(e))

    async def get_report(
        self, report_id: uuid.UUID, user_id: str
    ) -> ReportResponse | None:
        report = await self._repo.get_by_id(report_id)
        if not report or report.user_id != user_id:
            return None

        download_url = None
        if report.status == "completed" and report.file_key:
            download_url = await self._minio.get_download_url(
                report.file_key, self._download_url_expiry
            )
        return self._to_response(report, download_url)

    async def download_report(
        self, report_id: uuid.UUID, user_id: str
    ) -> tuple[BytesIO, str] | None:
        report = await self._repo.get_by_id(report_id)
        if (
            not report
            or report.user_id != user_id
            or report.status != "completed"
            or not report.file_key
        ):
            return None

        file_content = await self._minio.get_file(report.file_key)
        return file_content, report.file_name

    async def list_reports(
        self, user_id: str, limit: int = 50, offset: int = 0
    ) -> ReportListResponse:
        reports, total = await self._repo.list_by_user(user_id, limit, offset)
        return ReportListResponse(
            reports=[self._to_response(r) for r in reports],
            total=total,
        )

    async def delete_report(
        self, report_id: uuid.UUID, user_id: str
    ) -> bool:
        report = await self._repo.get_by_id(report_id)
        if not report or report.user_id != user_id:
            return False

        file_key = await self._repo.delete_report(report_id)
        if file_key:
            try:
                await self._minio.delete_file(file_key)
            except Exception as e:
                logger.warning(f"Ошибка при удалении из MinIO: {e}")
        return True

    async def _fetch_data(self, request: ReportRequest) -> dict:
        platform = request.platform.value
        account_id = request.account_id
        df = request.date_from
        dt = request.date_to

        if request.report_type == ReportType.full:
            results = await asyncio.gather(
                self._analytics.get_overview(platform, account_id),
                self._analytics.get_followers(platform, account_id, df, dt),
                self._analytics.get_posts(platform, account_id, df, dt),
                self._analytics.get_engagement(platform, account_id, df, dt),
                self._analytics.get_growth(
                    platform, account_id, df, dt, request.projection_days
                ),
                self._analytics.get_content_performance(
                    platform, account_id, df, dt
                ),
                self._analytics.get_posting_patterns(
                    platform, account_id, df, dt
                ),
                self._analytics.get_trends(platform, account_id, df, dt),
                return_exceptions=True,
            )
            keys = [
                "overview",
                "followers",
                "posts",
                "engagement",
                "growth",
                "content_performance",
                "posting_patterns",
                "trends",
            ]
            data = {}
            for key, result in zip(keys, results):
                if isinstance(result, Exception):
                    logger.warning(f"Ошибка получения файла {key}: {result}")
                else:
                    data[key] = result
            return data

        data: dict = {}
        if request.report_type == ReportType.overview:
            data["overview"] = await self._analytics.get_overview(
                platform, account_id
            )
        elif request.report_type == ReportType.engagement:
            data["engagement"] = await self._analytics.get_engagement(
                platform, account_id, df, dt
            )
        elif request.report_type == ReportType.growth:
            data["growth"] = await self._analytics.get_growth(
                platform, account_id, df, dt, request.projection_days
            )
        elif request.report_type == ReportType.content:
            data["content_performance"] = (
                await self._analytics.get_content_performance(
                    platform, account_id, df, dt
                )
            )
        return data

    @staticmethod
    def _to_response(
        report, download_url: str | None = None
    ) -> ReportResponse:
        return ReportResponse(
            id=report.id,
            user_id=report.user_id,
            platform=report.platform,
            account_id=report.account_id,
            report_type=report.report_type,
            status=report.status,
            file_name=report.file_name,
            download_url=download_url,
            error=report.error,
            created_at=report.created_at,
            completed_at=report.completed_at,
        )
