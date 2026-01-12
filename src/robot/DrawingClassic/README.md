# DrawingClassic

## Tóm tắt
- Mục đích: [Mô tả: ví dụ xử lý bản vẽ cổ điển, chuẩn hoá và gửi kết quả].

## Flow (quy trình)
1. Nhận input (file bản vẽ, params).
2. Dùng automation để truy cập nguồn (SharePoint/web).
3. Chạy xử lý (chuyển đổi, chỉnh sửa file, chạy macro nếu cần).
4. Upload kết quả, cập nhật log và trả về báo cáo.

## Các file cần thiết
- `tasks.py` — task chính
- `automation/` — wrappers/điều khiển browser hoặc I/O
- `__init__.py`

---

## Chạy & Debug
- Gọi task qua Celery hoặc chạy trực tiếp để debug.

## Ghi chú
- Thêm thông tin chi tiết (tên function, các biến env) nếu cần cho maintainers.
