# MejiIrisumiCheck

## Tóm tắt
- Mục đích: [Kiểm tra, validate dữ liệu Meji/Irisumi].

## Flow (quy trình)
1. Nhận input (file/danh sách).
2. Dùng `automation/` để lấy dữ liệu nếu cần.
3. Chạy kiểm tra/validate, xuất báo cáo lỗi.
4. Upload kết quả.

## Các file cần thiết
- `tasks.py`
- `automation/`

---

## Chạy & Debug
- Test các hàm validate nhỏ riêng lẻ trước khi chạy toàn bộ task qua Celery.
