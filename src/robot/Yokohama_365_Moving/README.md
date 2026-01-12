# Yokohama_365_Moving

## Tóm tắt
- Mục đích: [Di chuyển thư mục trên SharePoint/Sharepoint folder moving tự động].

## Flow (quy trình)
1. Nhập input (thư mục nguồn / đích / params).
2. Quản lý access token (xem `Access_token/`).
3. Chạy logic chính `tasks.py` để di chuyển folder hoặc tệp.
4. Log kết quả trong `bot_log.log`.

## Các file cần thiết
- `tasks.py`
- `Access_token/` — chứa token/credential handling
- `sharepoint_folder_moving_V1_1.exe` — helper exe
- `bot_log.log` — log file

---

## Chạy & Debug
- Kiểm tra `Access_token/` để đảm bảo credentials hợp lệ; xem `bot_log.log` để debug.
