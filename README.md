# RPA Control Center

## 🛠 Công nghệ sử dụng

* **Core:** Python (FastAPI)
* **Task Queue:** Celery & Redis
* **Database:** MySQL (Alembic cho migration)
* **Real-time:** Socket.IO

## 🚀 Cài đặt & Chạy dự án

### 1. Yêu cầu tiên quyết
Đảm bảo máy tính của bạn đã cài đặt:
* [Python 3.10+](https://www.python.org/)
* [Docker](https://www.docker.com/)
* [uv](https://github.com/astral-sh/uv)

### 2. Thiết lập môi trường
Copy file cấu hình mẫu và điền thông tin của bạn vào:
```bash
cp .env.example .env
