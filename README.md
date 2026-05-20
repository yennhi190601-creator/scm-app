# SCM Request Management System
# Version 1 -- Backup because vesrion 2 have database errors
## Cấu trúc thư mục 

```
scm_app/
├── app.exe           ← Chạy file này để khởi động hệ thống
├── config.json       
├── frontend/
│   └── index.html    
└── data/
    └── scm_data.db   ← Database (backup định kỳ)
```

---

## Hướng dẫn sử dụng (dành cho công ty)

### 1. Khởi động hệ thống
- Double-click vào file `app.exe`
- Mở trình duyệt, truy cập: http://localhost:8000
- Giữ cửa sổ cmd/terminal mở trong quá trình sử dụng

### 2. Tùy chỉnh danh sách dropdown — sửa file `config.json`

Mở `config.json` bằng Notepad, sửa các mục cần thiết:

```json
{
  "app_title": "Tiêu đề trang web",
  "company_name": "Tên phòng ban",

  "muc_dich_options": ["Nghiên cứu mới", "Thay thế vật tư", ...],
  "phan_loai_options": ["Vật tư A", "Linh kiện", ...],
  "phong_ban_options": ["Phòng SCM", "Phòng R&D", ...],
  "dvt_options": ["Cái", "Bộ", "Kg", ...],
  "pic_mua_hang_options": ["Nguyễn Văn A", ...]
}
```

⚠️ Lưu ý khi sửa config.json:
- Lưu file với encoding UTF-8
- Không xóa dấu ngoặc `[]` hay `{}`
- Mỗi mục cách nhau bằng dấu phẩy `,`
- Sau khi sửa, tắt và khởi động lại app.exe

### 3. Sửa giao diện — file `frontend/index.html`
- Mở bằng Notepad hoặc VS Code
- Chỉnh màu sắc, font chữ, tiêu đề cột trong phần `<style>`
- Không xóa các `id="..."` quan trọng

### 4. Backup dữ liệu
- Copy file `data/scm_data.db` ra nơi khác định kỳ
- Hoặc dùng nút "Xuất Excel" trên giao diện

---

## Tính năng chính

| Tính năng | Mô tả |
|---|---|
| Nhập liệu hàng loạt | Điền nhiều dòng cùng lúc, nhấn "Lưu tất cả" |
| Tìm kiếm | Tìm theo mã YC, tên vật tư, PIC |
| Duyệt trạng thái | Click vào badge "Chưa duyệt" / "Đã duyệt" để đổi |
| Xuất Excel | Xuất toàn bộ data có định dạng, màu sắc |
| Phân trang | 50 bản ghi / trang |

---

## Hướng dẫn compile (dành cho SCM data team — giữ nội bộ)

```bash
# Cài dependencies
pip install -r requirements.txt
pip install pyinstaller

# Compile thành exe
pyinstaller --onefile --add-data "frontend;frontend" --add-data "config.json;." app.py

# Output: dist/app.exe
```
WEB: https://scm-app-9q8y.onrender.com
