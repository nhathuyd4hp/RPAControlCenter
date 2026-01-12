# SeikyuOnline

## Tóm tắt
- Mục đích: [Xử lý Seikyu online—thu thập, kiểm tra, gửi hoá đơn online].

## Flow (quy trình)
1. Nhập input hoặc truy vấn API.
2. Dùng `automation/` để tương tác khi cần.
3. Xử lý và đẩy kết quả.

## Các file cần thiết
- `tasks.py`
- `api/`
- `automation/`

---

## Chạy & Debug
- Test từng phần xử lý trước khi chạy task toàn bộ via Celery.
