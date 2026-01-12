# ShukoTaoSo

## Tóm tắt
- Mục đích: [Tạo số/Shuko — xử lý tạo bản ghi và upload tài liệu].

## Flow (quy trình)
1. Nhập input (excel/danh sách).
2. Dùng `automation/` để truy cập hệ thống (nếu cần).
3. Chạy logic trong `tasks.py` và export file.
4. Upload kết quả và log.

## Các file cần thiết
- `tasks.py`
- `automation/`

---

## Chạy & Debug
- Test các hàm helper trước; dùng Celery để chạy task production.
