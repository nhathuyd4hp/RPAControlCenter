import shutil
import subprocess
from pathlib import Path

from celery import shared_task
from celery.app.task import Task

from src.core.config import settings
from src.service import StorageService as minio


@shared_task(bind=True, name="Yoko Mitsumori Waritsukezu Soufu")
def main(self: Task):
    exe_path = (
        Path(__file__).resolve().parents[2] / "robot" / "YokoMitsumoriWaritsukezuSoufu" / "図面・見積書送付_V1.8.exe"
    )
    cwd_path = exe_path.parent

    log_dir = Path(__file__).resolve().parents[3] / "logs"
    log_dir.mkdir(parents=True, exist_ok=True)
    log_file = log_dir / f"{self.request.id}.log"

    with open(log_file, "w", encoding="utf-8", errors="ignore") as f:
        process = subprocess.Popen([str(exe_path)], cwd=str(cwd_path), stdout=f, stderr=subprocess.STDOUT, text=True)
        process.wait()

    for path in [
        cwd_path / "Access_token_log",
        cwd_path / "Access_token",
        cwd_path / "Ankens",
        cwd_path / "Logs",
        cwd_path / "横浜",
    ]:
        if path.is_dir():
            shutil.rmtree(path, ignore_errors=True)
        if path.is_file():
            path.unlink(missing_ok=True)

    ProgressReports = cwd_path / "ProgressReports"
    zip_path = cwd_path / "ProgressReports.zip"

    shutil.make_archive(
        base_name=str(ProgressReports),
        format="zip",
        root_dir=str(cwd_path),
        base_dir="ProgressReports",
    )

    result = minio.fput_object(
        bucket_name=settings.RESULT_BUCKET,
        object_name=f"YokoMitsumoriWaritsukezuSoufu/{self.request.id}/ProgressReports.zip",
        file_path=str(zip_path),
        content_type="application/zip",
    )

    zip_path.unlink(missing_ok=True)

    return f"{settings.RESULT_BUCKET}/{result.object_name}"
