from src.core.playwright_patch import apply_patch

apply_patch()

from celery import Celery  # noqa
from src.core.config import settings  # noqa
from src.robot import *  # noqa
from src.worker_signals import *  # noqa

Worker = Celery(
    "orchestration",
    broker=settings.REDIS_CONNECTION_STRING,
    backend=settings.REDIS_CONNECTION_STRING,
)
