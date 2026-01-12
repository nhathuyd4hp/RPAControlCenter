import os
import subprocess
from pathlib import Path

from celery import shared_task


@shared_task(bind=True, name="Gửi Bản Vẽ Shuko")
def GuiBanVeShuko(self):
    exe_path = Path(__file__).resolve().parents[2] / "robot" / "GuiBanVeShuko" / "GuiBanVeShuko.exe"
    cwd_path = exe_path.parent

    log_dir = Path(__file__).resolve().parents[3] / "logs"
    log_dir.mkdir(parents=True, exist_ok=True)
    log_file = log_dir / f"{self.request.id}.log"

    env = os.environ.copy()
    env["PYTHONUNBUFFERED"] = "1"

    with open(log_file, "a", encoding="utf-8") as f:
        try:
            process = subprocess.Popen(
                [str(exe_path)],
                cwd=str(cwd_path),
                stdout=subprocess.PIPE,
                stderr=subprocess.STDOUT,
                stdin=subprocess.DEVNULL,
                env=env,
                shell=True,
            )

            for line in iter(process.stdout.readline, b""):
                try:
                    decoded_line = line.decode("utf-8", errors="replace")
                except Exception:
                    decoded_line = str(line)

                f.write(decoded_line)
                f.flush()

            process.wait()

        except Exception as e:
            f.write(f"\n[CRITICAL ERROR] Python subprocess failed: {str(e)}\n")
