from typing import Any
from taskiq_redis import RedisAsyncResultBackend, ListQueueBroker
from src.settings import settings

result_backend = RedisAsyncResultBackend(
    redis_url=settings.redis.url,
)

broker = ListQueueBroker(
    url=settings.redis.url,
    result_backend=result_backend,
).with_result_backend(result_backend)


@broker.on_event("startup")
async def startup(state: Any) -> None:
    from src.interfaces.api.containers import Container
    
    container = Container()
    container.wiring_config.modules.append("src.infrastructure.tasks.reports")
    await container.init_resources()
    state.container = container


@broker.on_event("shutdown")
async def shutdown(state: Any) -> None:
    if hasattr(state, "container"):
        await state.container.shutdown_resources()
