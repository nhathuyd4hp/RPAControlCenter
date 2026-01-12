# GuiMailNoukiKakunin

## Tóm tắt
- Mục đích: [Xác nhận ngày giao hàng thông qua mail tự động / tải mail và check trạng thái].

## Flow (quy trình)
1. Nhận input (danh sách order hoặc ngày).
2. Dùng `api/` hoặc `automation/` để lấy mail/chi tiết.
3. Xử lý mailbox, trích xuất thông tin, update trạng thái.
4. Ghi log và gửi thông báo nếu cần.

## Các file cần thiết
- `tasks.py`
- `api/` — endpoints helpers
- `automation/` — nếu cần chạy automation

---

## Chạy & Debug
- Chạy trực tiếp hoặc qua worker Celery; kiểm tra logs trong `__pycache__` khi dev.

## Ghi chú
- Cẩn thận với dữ liệu nhạy cảm trong mail; tránh commit credential.
