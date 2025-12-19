from datetime import datetime

from celery import shared_task


def from_date_default() -> datetime:
    now = datetime.now()
    if now.month == 1:
        return now.replace(year=now.year - 1, month=12, day=21)
    return now.replace(month=now.month - 1, day=21)


def to_date_default() -> datetime:
    now = datetime.now()
    return now.replace(day=20)


@shared_task(
    bind=True,
    name="Sakura",
)
def Sakura(
    self,
    from_date: datetime | str = from_date_default(),
    to_date: datetime | str = to_date_default(),
):
    pass
