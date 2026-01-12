from celery import Task

class InactiveTask(Task):
    active: bool = False

    def __call__(self, *args, **kwargs):
        if self.active is not True:
            raise RuntimeError("Inactive Task")
        return super().__call__(*args, **kwargs)