# KEIAIANKENKA

## Tóm tắt
- Mục đích: [Mô tả ngắn — ví dụ: tích hợp với KISTAR/Keiaistar để xử lý ankenka].

## Flow (quy trình)
1. Nhập input và/hoặc file Excel `Keiaistar.xlsx`.
2. Chạy logic trong `tasks.py`.
3. Nếu cần, tương tác với `KISTAR_AnkenkaV1.9.exe` để sinh output.
4. Upload kết quả/ghi log.

## Các file cần thiết
- `tasks.py`
- `Keiaistar.xlsx` (input template)
- `KISTAR_AnkenkaV1.9.exe` (executable helper)

---

## Chạy & Debug
- Kiểm tra compatibility khi gọi `.exe` từ Windows environment; chạy thử local.

## Ghi chú
- Lưu ý bản quyền và permission khi giữ file `.exe` trong repo.
