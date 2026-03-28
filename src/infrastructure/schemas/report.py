import uuid
from datetime import datetime
from enum import Enum

from pydantic import BaseModel, Field


class Platform(str, Enum):
    instagram = "instagram"
    tiktok = "tiktok"
    youtube = "youtube"


class ReportType(str, Enum):
    full = "full"
    overview = "overview"
    growth = "growth"
    engagement = "engagement"
    content = "content"


class ReportRequest(BaseModel):
    platform: Platform
    account_id: str
    report_type: ReportType = ReportType.full
    date_from: datetime | None = None
    date_to: datetime | None = None
    projection_days: int = Field(default=14, ge=1, le=90)


class ReportResponse(BaseModel):
    id: uuid.UUID
    user_id: str
    platform: str
    account_id: str
    report_type: str
    status: str
    file_name: str | None = None
    download_url: str | None = None
    error: str | None = None
    created_at: datetime
    completed_at: datetime | None = None


class ReportListResponse(BaseModel):
    reports: list[ReportResponse]
    total: int
