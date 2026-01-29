import shutil
import subprocess
from pathlib import Path

from celery import shared_task
from celery.app.task import Task


@shared_task(bind=True)
def main(self: Task):
    id: str = self.request.id
    exe_path = Path(__file__).resolve().parents[2] / "robot" / "TamahomeNoukihikaku" / "タマ納期比較_V2.0.exe"
    log_dir = Path(__file__).resolve().parents[3] / "logs"
    log_dir.mkdir(parents=True, exist_ok=True)
    log_file = log_dir / f"{id}.log"

    with open(log_file, "w", encoding="utf-8", errors="ignore") as f:
        process = subprocess.Popen(
            [str(exe_path)], cwd=str(exe_path.parent), stdout=f, stderr=subprocess.STDOUT, text=True
        )
        process.wait()

    for path in [
        exe_path.parent / "Access_token_log",
        exe_path.parent / "Logs",
        exe_path.parent / "Access_token",
    ]:
        if path.is_file():
            path.unlink(missing_ok=True)
        if path.is_dir():
            shutil.rmtree(path, ignore_errors=True)
