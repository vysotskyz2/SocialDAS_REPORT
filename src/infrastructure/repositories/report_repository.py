import uuid
from datetime import datetime
from sqlalchemy import insert, select, update, delete, func
from sqlalchemy.ext.asyncio import AsyncSession
from src.infrastructure.models.report import Report


class ReportRepository:
    def __init__(self, session: AsyncSession) -> None:
        self._session = session

    async def create(
        self,
        user_id: str,
        platform: str,
        account_id: str,
        report_type: str,
    ) -> Report:
        result = await self._session.execute(
            insert(Report)
            .values(
                user_id=user_id,
                platform=platform,
                account_id=account_id,
                report_type=report_type,
            )
            .returning(Report)
        )
        return result.scalars().one()

    async def get_by_id(self, report_id: uuid.UUID) -> Report | None:
        result = await self._session.execute(
            select(Report).where(Report.id == report_id)
        )
        return result.scalars().one_or_none()

    async def list_by_user(
        self, user_id: str, limit: int = 50, offset: int = 0
    ) -> tuple[list[Report], int]:
        count_result = await self._session.execute(
            select(func.count())
            .select_from(Report)
            .where(Report.user_id == user_id)
        )
        total = count_result.scalar()

        result = await self._session.execute(
            select(Report)
            .where(Report.user_id == user_id)
            .order_by(Report.created_at.desc())
            .limit(limit)
            .offset(offset)
        )
        return list(result.scalars().all()), total

    async def update_completed(
        self, report_id: uuid.UUID, file_key: str, file_name: str
    ) -> None:
        await self._session.execute(
            update(Report)
            .where(Report.id == report_id)
            .values(
                status="completed",
                file_key=file_key,
                file_name=file_name,
                completed_at=datetime.utcnow(),
            )
        )

    async def update_failed(
        self, report_id: uuid.UUID, error: str
    ) -> None:
        await self._session.execute(
            update(Report)
            .where(Report.id == report_id)
            .values(status="failed", error=error)
        )

    async def delete_report(self, report_id: uuid.UUID) -> str | None:
        result = await self._session.execute(
            select(Report.file_key).where(Report.id == report_id)
        )
        file_key = result.scalar_one_or_none()

        await self._session.execute(
            delete(Report).where(Report.id == report_id)
        )
        return file_key
