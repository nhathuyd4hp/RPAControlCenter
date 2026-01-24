from celery import shared_task
from celery.app.task import Context, Task


@shared_task(bind=True, name="Hajime Ankenka")
def HajimeAnkenka(self: Task):
    context: Context = self.request
    id = context.id
    return id
