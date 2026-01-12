# FuriwakeOsaka

## Tóm tắt
- Mục đích: [Mô tả ngắn về nhiệm vụ chính của robot].

## Flow (quy trình)
1. Nhận input (csv/excel/param).
2. Xác thực/Quản lý token (xem `token_manager.py`).
3. Tải dữ liệu và xử lý (xử lý BOM hoặc dữ liệu theo logic riêng).
4. Lưu kết quả vào `Results/` và ghi log tại `Logs/`.

## Các file cần thiết (đã phát hiện)
- `tasks.py`
- `Main.py` — entry script (nếu cần chạy độc lập)
- `token_manager.py`, `config.py`, `config_access_token.py` — config & token
- `bom_downloader.py` — helper download
- `Logs/`, `Results/`

---

## Chạy & Debug
- Chạy `Main.py` để thử local hoặc gọi các task qua Celery.

## Ghi chú
- Token/credentials được quản lý trong `token_manager.py` & các file config — bảo mật khi commit.
