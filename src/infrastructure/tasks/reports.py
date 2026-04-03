import uuid
import asyncio
from loguru import logger
from src.infrastructure.tasks.broker import broker
from src.application.services.report_service import ReportService
from src.infrastructure.schemas.report import ReportRequest
from src.infrastructure.models.base import async_session_maker
from src.interfaces.api.containers import db_session_context


@broker.task
async def generate_report_task(
    report_id: uuid.UUID,
    user_id: str,
    request_data: dict,
) -> None:
    from src.infrastructure.tasks.broker import container
    
    request = ReportRequest(**request_data)
    
    async with async_session_maker() as session:
        token = db_session_context.set(session)
        try:
            service_res = container.report_service()
            if asyncio.iscoroutine(service_res) or isinstance(service_res, asyncio.Future):
                service = await service_res
            else:
                service = service_res
            
            await service.process_report_generation(report_id, user_id, request)
            await session.commit()
        except Exception as e:
            await session.rollback()
            logger.error(f"Ошибка при обработке отчета {report_id}: {e}")
        finally:
            db_session_context.reset(token)
