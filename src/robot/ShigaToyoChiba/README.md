# ShigaToyoChiba

## Tóm tắt
- Mục đích: [Ví dụ: thu thập dữ liệu/điều phối bản vẽ, gửi báo cáo cho Shiga/Toyo/Chiba].

## Flow (quy trình)
1. Nhận input (ngày, range, param).
2. Dùng `automation/` và `api/` để truy xuất/ghi nhận dữ liệu.
3. Xử lý chính trong `tasks.py`.
4. Upload kết quả, ghi log, gửi thông báo nếu cần.

## Các file cần thiết (đã phát hiện)
- `tasks.py`
- `api/` — helper endpoint wrappers
- `automation/` — scripts/helpers để tương tác

---

## Chạy & Debug
- Gọi task chính qua Celery: `from src.robot.ShigaToyoChiba import tasks; tasks.<main_task>.delay(...)`.
- Để debug local: import và gọi trực tiếp hàm xử lý với dữ liệu mẫu.

## Ghi chú
- Cập nhật README khi thêm file cấu hình hoặc thay đổi flow.
