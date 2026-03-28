from contextlib import asynccontextmanager

from fastapi import FastAPI, Request, Depends, status, HTTPException
from sqlalchemy import text
from sqlalchemy.ext.asyncio import AsyncSession
from dependency_injector.wiring import inject, Provide

from src.infrastructure.models.base import async_session_maker
from src.interfaces.api.containers import Container, db_session_context
from src.interfaces.api.routers import reports


@asynccontextmanager
async def lifespan(app: FastAPI):
    container = Container()
    container.wire()
    await container.init_resources()

    minio_client = container.minio_client()
    await minio_client.ensure_bucket()

    yield

    await container.shutdown_resources()


app = FastAPI(title="SocialDAS Report Service", lifespan=lifespan)


@app.middleware("http")
async def db_session_middleware(request: Request, call_next):
    async with async_session_maker() as session:
        token = db_session_context.set(session)
        try:
            response = await call_next(request)
            await session.commit()
            return response
        except Exception as e:
            await session.rollback()
            raise e
        finally:
            db_session_context.reset(token)


@app.get("/health", status_code=status.HTTP_200_OK)
@inject
async def health_check(
    db: AsyncSession = Depends(Provide[Container.db_session]),
):
    try:
        await db.execute(text("SELECT 1"))
        return {"status": "healthy", "database": "connected"}
    except Exception:
        raise HTTPException(
            status_code=status.HTTP_503_SERVICE_UNAVAILABLE,
            detail="Database unavailable",
        )


app.include_router(reports.router)
