# KyushuOsaka

## Tóm tắt
- Mục đích: [Mô tả nhiệm vụ robot KyushuOsaka].

## Flow (quy trình)
1. Input nhận từ param hoặc API.
2. Sử dụng `automation/` để tương tác với web/service.
3. Xử lý dữ liệu và upload.

## Các file cần thiết
- `tasks.py`
- `api/` — helper endpoint wrappers
- `automation/`

---

## Chạy & Debug
- Gọi task qua Celery hoặc run trực tiếp các hàm để debug.
