# OHDMITSUMORI

## Tóm tắt
- Mục đích: [Xử lý các tác vụ liên quan OHD Mitsumori].

## Flow (quy trình)
1. Input: file/params.
2. Chạy logic trong `tasks.py` (có thể gọi `OHDNew.exe`).
3. Sinh báo cáo `OHD図面送付結果.xlsx` hoặc tương tự.
4. Upload & log.

## Các file cần thiết
- `tasks.py`
- `OHDNew.exe` — executable hỗ trợ
- `OHD図面送付結果.xlsx` — mẫu/report

---

## Chạy & Debug
- Kiểm tra trước khi gọi exe; chạy local để debug file output.
