# GuiBanVeToei

## Tóm tắt
- Mục đích: [Gửi bản vẽ, xử lý xác nhận và upload kết quả].

## Flow (quy trình)
1. Input: file bản vẽ / metadata.
2. Dùng `automation/` để tương tác site.
3. Xử lý và tạo output.
4. Upload kết quả và báo trạng thái.

## Các file cần thiết
- `tasks.py`
- `automation/`
- `__init__.py`

---

## Chạy & Debug
- Gọi task chính qua Celery hoặc import hàm để chạy local.

## Ghi chú
- Thêm chi tiết flow khi công việc có nhiều bước phụ.
