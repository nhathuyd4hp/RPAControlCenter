import shutil
import subprocess
from pathlib import Path
from celery import shared_task

@shared_task(bind=True,name="Keiai Nouki Irai")
def keiai_nouki_irai(self):
    exe_path = Path(__file__).resolve().parents[2] / "robot" / "KeiaiNoukiIrai" / "KISTAR_Nouki_V2_FItUkGg.5.exe"
    cwd_path = exe_path.parent

    log_dir = Path(__file__).resolve().parents[3] / "logs"
    log_dir.mkdir(parents=True, exist_ok=True)
    log_file = log_dir / f"{self.request.id}.log"

    with open(log_file, "w", encoding="utf-8", errors="ignore") as f:
        process = subprocess.Popen([str(exe_path)], cwd=str(cwd_path), stdout=f, stderr=subprocess.STDOUT, text=True)
        process.wait()

    # Clean
    log_file = cwd_path / "Access_token_log"
    try:
        log_file.unlink()
    except Exception:
        pass
    
    logs_folder = cwd_path / "Logs"
    access_token_folder = cwd_path / "Access_token"
    for path in (logs_folder, access_token_folder):
        shutil.rmtree(path, ignore_errors=True)