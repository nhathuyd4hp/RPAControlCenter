# Tochigi

## Tóm tắt
- Mục đích: [Xử lý tác vụ Tochigi—thu thập, xử lý, gửi báo cáo].

## Flow (quy trình)
1. Nhận params/input.
2. Sử dụng `api/` và `automation/` để thu thập dữ liệu.
3. Xử lý trong `tasks.py`.
4. Upload/ghi log.

## Các file cần thiết
- `tasks.py`
- `api/`
- `automation/`

---

## Chạy & Debug
- Test hàm nhỏ riêng lẻ, chạy Celery cho production.
