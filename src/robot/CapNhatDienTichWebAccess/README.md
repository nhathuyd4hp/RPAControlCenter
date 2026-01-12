# CapNhatDienTichWebAccess

## Tóm tắt
- Mục đích: [Một câu mô tả nhiệm vụ chính — ví dụ: cập nhật diện tích/ghi nhận dữ liệu từ web].

## Flow (quy trình)
1. Nhận input (thời gian/ID/danh sách).
2. Đăng nhập — lấy dữ liệu từ web hoặc API.
3. Tiền xử lý (clean, parse).
4. Ghi/đẩy dữ liệu (MinIO / database / SharePoint).

## Các file cần thiết (tự động phát hiện)
- `tasks.py` (bắt buộc)
- `automation/` (helper scripts)
- `__init__.py`

---

## Chạy & Debug
- Sử dụng Celery: import task và gọi `.delay()` hoặc chạy trực tiếp hàm để debug.

## Ghi chú
- Nếu có file cấu hình đặc thù (ví dụ khóa truy cập), lưu giữ trong config an toàn.
