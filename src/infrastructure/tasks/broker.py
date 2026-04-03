from typing import Any
from taskiq_redis import RedisAsyncResultBackend, ListQueueBroker
from src.settings import settings
from src.interfaces.api.containers import Container

result_backend = RedisAsyncResultBackend(
    redis_url=settings.redis.url,
)

broker = ListQueueBroker(
    url=settings.redis.url,
).with_result_backend(result_backend)

container = Container()


@broker.on_event("startup")
async def startup(state: Any) -> None:
    import src.infrastructure.tasks.reports as reports_tasks
    container.wire(modules=[reports_tasks])
    await container.init_resources()
    state.container = container


@broker.on_event("shutdown")
async def shutdown(state: Any) -> None:
    await container.shutdown_resources()
