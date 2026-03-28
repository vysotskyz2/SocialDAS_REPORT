import uuid
from loguru import logger
from dependency_injector.wiring import inject, Provide
from src.infrastructure.tasks.broker import broker
from src.interfaces.api.containers import Container
from src.application.services.report_service import ReportService
from src.infrastructure.schemas.report import ReportRequest


@broker.task
@inject
async def generate_report_task(
    report_id: uuid.UUID,
    user_id: str,
    request_data: dict,
    service: ReportService = Provide[Container.report_service],
) -> None:
    logger.info(f"Обработка отчета: {report_id}")

    request = ReportRequest(**request_data)
    
    try:
        await service.process_report_generation(report_id, user_id, request)
    except Exception as e:
        logger.error(f"Ошибка при обработке отчета: {report_id} - {e}")
