# AndPadShuko

## Tóm tắt
- Mục đích: [Viết một câu ngắn mô tả nhiệm vụ chính của robot AndPadShuko].

## Flow (quy trình)
1. Nhận input (ví dụ: ngày, file cấu hình).
2. Đăng nhập/Ủy quyền tới dịch vụ cần thiết (SharePoint / web).
3. Tải dữ liệu (download) hoặc đọc file nguồn.
4. Xử lý dữ liệu (sửa, chuyển đổi, chạy macro nếu cần).
5. Tải kết quả lên (MinIO / SharePoint) và ghi log kết quả.

## Các file cần thiết
- `tasks.py` — chứa Celery tasks của robot.
- `automation/` — helper, script tự động (Playwright/SharePoint wrappers...).
- `__init__.py`

---

## Chạy & Debug
- Chạy worker Celery và enqueue task: `from src.robot.AndPadShuko import tasks; tasks.<main_task>.delay(...)`.
- Để debug local, import hàm chính và chạy trực tiếp với dữ liệu mock.

## Ghi chú
- Cập nhật các phần Tóm tắt/Flow/Files khi thay đổi behavior hoặc file phụ thuộc.
