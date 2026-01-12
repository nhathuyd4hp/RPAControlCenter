# Kaneka

## Tóm tắt
- Mục đích: [Mô tả nhiệm vụ robot Kaneka].

## Flow (quy trình)
1. Nhận input/requests.
2. Chạy logic chính trong `tasks.py` (dùng `common/`, `core/`, `service/` nếu cần).
3. Tạo và lưu kết quả, upload nếu cần.

## Các file cần thiết
- `tasks.py`
- `common/`, `core/`, `service/` — helpers & abstrations
- `__init__.py`

---

## Chạy & Debug
- Import task và chạy trực tiếp hoặc enqueue qua Celery.

## Ghi chú
- Khi refactor helpers, cập nhật phần "Các file cần thiết" để giữ docs đúng.
