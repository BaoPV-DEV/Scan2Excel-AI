# Tab 2: Tách File Nhân Viên Theo Tổ

## 📋 Mô Tả Chức Năng

Tab này đọc danh sách nhân viên gốc từ 1 file Excel chứa toàn bộ nhân viên của nhà máy, sau đó tự động tách thành từng file Excel riêng biệt cho từng tổ/bộ phận. Mỗi file tách sử dụng template chuẩn để đảm bảo định dạng đồng nhất.

---

## 🎯 Quy Trình Làm Việc

### 1. **Chọn File Danh Sách Nhân Viên Gốc**
- Nhấn **"📂 Chọn File"** trong phần "Chọn File Danh Sách NV gốc (.xlsx)"
- Chọn file Excel chứa danh sách toàn bộ nhân viên
- File này phải có cột **"Tổ"** và **"Mã nhân viên"** để ứng dụng có thể phân loại

### 2. **Cấu Hình Thời Gian Báo Cáo**
- Chọn **Tháng** (01 → 12) 
- Chọn **Năm** (thường là năm hiện tại)
- Thông tin này sẽ được sử dụng để:
  - Tạo đường dẫn lưu trữ output: `D:\Linh_Salary_Tool\01_danh_sach_chia_to\YYYY\MM\`
  - Ghi vào metadata để Tab 5 có thể đọc

### 3. **Chọn Folder Output (Tuỳ Chọn)**
- Nhấn **"📂 Chọn Folder"** nếu muốn lưu output vào vị trí khác
- Mặc định: `D:\Linh_Salary_Tool\01_danh_sach_chia_to\YYYY\MM\`

### 4. **Bắt Đầu Tách File**
- Nhấn nút **"✂️ TÁCH FILE"** (nút xanh)
- Ứng dụng sẽ:
  - **Đọc** danh sách nhân viên từ file gốc
  - **Phân loại** nhân viên theo cột "Tổ"
  - **Tạo file mới** cho mỗi tổ sử dụng template chuẩn
  - **Điền dữ liệu** nhân viên vào sheet "Danh sách"
  - **Lưu metadata.json** để Tab 5 có thể đọc

### 5. **Theo Dõi Tiến Độ**
- Thanh tiến độ hiển thị % hoàn tất
- Logs chi tiết hiển thị từng bước:
  - Số nhân viên đã xử lý
  - Tên các file tách được tạo
  - Số tổ được phát hiện
  - Thông tin lưu metadata

---

## 📁 File Sử Dụng

| Loại | File | Vị Trí |
|------|------|--------|
| **Input** | Danh sách nhân viên gốc | Bất kỳ (do user chọn) |
| **Template** | to_may_template.xlsx, kiem_hoa_template.xlsx | `Template/sx/`, `Template/gt/` |
| **Output** | Các file tách theo tổ | `D:\Linh_Salary_Tool\01_danh_sach_chia_to\YYYY\MM\` |
| **Metadata** | metadata.json | `D:\Linh_Salary_Tool\04_data\YYYY\MM\` |

---

## 📊 Cấu Trúc Output

```
D:\Linh_Salary_Tool\01_danh_sach_chia_to\2026\04\
├── sx/
│   ├── Tổ Cắt - 04.xlsx
│   ├── Tổ May - 04.xlsx
│   └── ...
├── gt/
│   ├── Bảo vệ - 04.xlsx
│   ├── Kiểm hoá - 04.xlsx
│   └── ...
```

---

## 📄 Cấu Trúc File Tách

Mỗi file tách có cấu trúc:
- **Sheet 1: "Danh sách"** → Danh sách nhân viên của tổ
- **Sheet 2: "Bang TH nop"** → Bảng tính lương (điền sau ở Tab 4)
- **Cột B**: Tên nhân viên
- **Cột C**: Mã nhân viên
- **Cột D-AC**: Các tổ/bộ phận (tùy theo cấu hình Tab 1)

---

## 🔑 Cột Input File Gốc (Bắt Buộc)

| Cột | Mục Đích |
|-----|----------|
| Mã NV / Mã nhân viên | Định danh duy nhất cho nhân viên |
| Tên / Họ tên | Tên đầy đủ của nhân viên |
| Tổ / Bộ phận | Tuyên bố nhân viên thuộc tổ nào |

---

## ⚠️ Lưu Ý Quan Trọng

- **File gốc không được thay đổi**: Ứng dụng chỉ đọc, không sửa file gốc
- **Template phải tồn tại**: Phải chạy Tab 1 trước để cấu hình template
- **Cột phân loại**: File gốc phải có cột "Tổ" hoặc "Bộ phận" (ứng dụng sẽ tự tìm)
- **Tự động phát hiện template**: Ứng dụng xác định loại template dựa trên nội dung file gốc
  - Nếu có "may" → dùng `to_may_template.xlsx`
  - Nếu có "bảo vệ", "kiểm hoá" → dùng `kiem_hoa_template.xlsx`

---

## 🔄 Luồng Dữ Liệu

```
File Danh Sách Nhân Viên Gốc
    ↓
[Đọc & Phân loại theo Tổ]
    ↓
Lấy Template Phù Hợp
    ↓
Điền Nhân Viên vào từng File
    ↓
Lưu Output (01_danh_sach_chia_to)
    ↓
Lưu Metadata (04_data) → Để Tab 5 đọc
```

---

## 💡 Ví Dụ Thực Tế

**File gốc có:**
- 50 nhân viên tổ Cắt
- 30 nhân viên tổ May
- 20 nhân viên tổ Kiểm hoá

**Output sẽ là:**
- `sx/Tổ Cắt - 04.xlsx` (50 nhân viên)
- `sx/Tổ May - 04.xlsx` (30 nhân viên)
- `gt/Kiểm hoá - 04.xlsx` (20 nhân viên)

Mỗi file sử dụng template chuẩn và sẵn sàng cho Tab 3 (Quét) và Tab 4 (Tích hợp)
