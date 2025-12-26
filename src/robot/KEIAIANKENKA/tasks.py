import subprocess
from pathlib import Path
from celery import shared_task

@shared_task(bind=True,name="KEIAI ANKENKA")
def KEIAI_ANKENKA(self):
    exe_path = Path(__file__).resolve().parents[2] / "robot" / "KEIAIANKENKA" / "KISTAR_AnkenkaV1.9.exe"

    log_dir = Path(__file__).resolve().parents[3] / "logs"
    log_dir.mkdir(parents=True, exist_ok=True)

    log_file = log_dir / f"{self.request.id}.log"

    with open(log_file, "w", encoding="utf-8", errors="ignore") as f:
        process = subprocess.Popen(
            [str(exe_path)], cwd=str(exe_path.parent), stdout=f, stderr=subprocess.STDOUT, text=True
        )
        process.wait()