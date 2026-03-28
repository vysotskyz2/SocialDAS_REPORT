from taskiq_redis import RedisAsyncResultBackend, RedisBroker
from taskiq.receiver import Receiver
from taskiq_dependencies import DependencyGraph
from src.settings import settings
from src.interfaces.api.containers import Container

result_backend = RedisAsyncResultBackend(
    redis_url=settings.redis.url,
)

broker = RedisBroker(
    url=settings.redis.url,
    result_backend=result_backend,
).with_result_backend(result_backend)

container = Container()
container.wiring_config.modules.append("src.infrastructure.tasks.reports")

@broker.on_event("startup")
async def startup() -> None:
    await container.init_resources()

@broker.on_event("shutdown")
async def shutdown() -> None:
    await container.shutdown_resources()
