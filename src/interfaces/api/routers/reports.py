import uuid
from fastapi import APIRouter, Depends, HTTPException, Query, status
from fastapi.responses import StreamingResponse
from dependency_injector.wiring import inject, Provide
from src.interfaces.api.containers import Container
from src.interfaces.api.dependencies.auth import get_current_user
from src.application.services.report_service import ReportService
from src.infrastructure.schemas.report import (
    ReportRequest,
    ReportResponse,
    ReportListResponse,
)

router = APIRouter(prefix="/api/v1/reports", tags=["reports"])


@router.post(
    "/generate",
    response_model=ReportResponse,
    status_code=status.HTTP_201_CREATED,
)
@inject
async def generate_report(
    request: ReportRequest,
    user_id: str = Depends(get_current_user),
    service: ReportService = Depends(Provide[Container.report_service]),
):
    return await service.generate_report(user_id, request)


@router.get("/", response_model=ReportListResponse)
@inject
async def list_reports(
    user_id: str = Depends(get_current_user),
    limit: int = Query(default=50, ge=1, le=100),
    offset: int = Query(default=0, ge=0),
    service: ReportService = Depends(Provide[Container.report_service]),
):
    return await service.list_reports(user_id, limit, offset)


@router.get("/{report_id}", response_model=ReportResponse)
@inject
async def get_report(
    report_id: uuid.UUID,
    user_id: str = Depends(get_current_user),
    service: ReportService = Depends(Provide[Container.report_service]),
):
    report = await service.get_report(report_id, user_id)
    if not report:
        raise HTTPException(
            status_code=status.HTTP_404_NOT_FOUND,
            detail="Отчет не найден",
        )
    return report


@router.get("/{report_id}/download")
@inject
async def download_report(
    report_id: uuid.UUID,
    user_id: str = Depends(get_current_user),
    service: ReportService = Depends(Provide[Container.report_service]),
):
    result = await service.download_report(report_id, user_id)
    if not result:
        raise HTTPException(
            status_code=status.HTTP_404_NOT_FOUND,
            detail="Report file not found",
        )

    file_content, file_name = result
    return StreamingResponse(
        file_content,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f"attachment; filename={file_name}"},
    )


@router.delete("/{report_id}", status_code=status.HTTP_204_NO_CONTENT)
@inject
async def delete_report(
    report_id: uuid.UUID,
    user_id: str = Depends(get_current_user),
    service: ReportService = Depends(Provide[Container.report_service]),
):
    deleted = await service.delete_report(report_id, user_id)
    if not deleted:
        raise HTTPException(
            status_code=status.HTTP_404_NOT_FOUND,
            detail="Report not found",
        )
