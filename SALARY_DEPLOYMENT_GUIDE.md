# 📊 Hệ Thống Triển Khai Lương Tập Trung (JSON Config)

## ✨ Đặc Điểm Chính

- ✅ **Config Tập Trung**: Tất cả template + source files định nghĩa trong 1 file JSON
- ✅ **Tự Động Mapping**: Input tháng/năm → Auto map đến `D:\Linh_Salary_Tool\{YYYY}\{MM}\`
- ✅ **Linh Hoạt**: Mỗi template, mỗi sheet, mỗi cột có thể khác nhau
- ✅ **Error Handling**: Nếu file/sheet chưa có → Để trống + Log cảnh báo
- ✅ **One-Click Deploy**: Nhấn 1 nút, tất cả tự động

---

## 🗂️ Cấu Trúc Thư Mục

```
D:\Linh_Salary_Tool\
├── 2026\
│   ├── 01\
│   │   ├── Danh sách CBCNV bản dùng làm lương T01.xlsx
│   │   ├── Chấm công 01.26.xlsx
│   │   ├── phụ cấp con nhỏ năm 2026.xlsx
│   │   └── ...
│   ├── 02\
│   └── ... (tháng khác)
└── 2027\ (năm khác)

Template\
├── salary_sources.json  ✨ CONFIG TẬP TRUNG (chỉnh sửa ở đây)
├── gt\
│   └── bep_an_va_cong_vu_template.xlsx
└── sx\
    └── ...

01_danh_sach_chia_to\ (Output từ Tab 1)
├── 2026\
│   └── 01\
│       ├── gt\
│       │   └── Bếp ăn - 01.xlsx
│       └── sx\
│           └── ...
```

---

## 📋 JSON Config Structure

File: `Template\salary_sources.json`

```json
{
  "base_path": "D:\\Linh_Salary_Tool",
  "templates": {
    "gt/bep_an_va_cong_vu_template.xlsx": {
      "sheets": {
        "Bang TH nop": {
          "start_row": 9,
          "row_step": 2,
          "columns": {
            "D": {
              "source_file": "Danh sách CBCNV bản dùng làm lương T{mm_int}.xlsx",
              "source_sheet": "Danh sách",
              "source_range": "$D$1:$AS$1000",
              "column_index": 42
            }
          }
        }
      }
    }
  }
}
```

### Placeholders Hỗ Trợ

| Placeholder | Ví dụ | Ghi Chú |
|---|---|---|
| `{yyyy}` | 2026 | Năm hiện tại |
| `{yy_short}` | 26 | 2 chữ số cuối năm |
| `{mm}` | 01 | Tháng (2 chữ số) |
| `{mm_int}` | 1 | Tháng (số) |
| `{prev_year}` | 2025 | Năm trước |
| `{prev_month}` | 12 | Tháng trước (2 chữ số) |
| `{prev_month_int}` | 12 | Tháng trước (số) |

---

## 🔄 Workflow

```
Step 1: Tab 1 - Tách File Nhân Viên
   Input: File gốc + tháng/năm
   Output: File split theo tổ (VD: Bếp ăn - 01.xlsx)
   Lưu tại: 01_danh_sach_chia_to\2026\01\gt\

Step 2: Chuẩn Bị Dữ Liệu Ngoài
   Đảm bảo file source sẵn sàng:
   - D:\Linh_Salary_Tool\2026\01\Danh sách CBCNV bản dùng làm lương T1.xlsx
   - D:\Linh_Salary_Tool\2026\01\Chấm công 01.26.xlsx
   - v.v.

Step 3: Tab 5 - Triển Khai Lương (ONE-CLICK)
   1. Chọn thư mục: 01_danh_sach_chia_to\2026\01\gt\
   2. Chọn template: gt/bep_an_va_cong_vu_template.xlsx
   3. Chọn tháng: 01
   4. Chọn năm: 2026
   5. Nhấn "TRIỂN KHAI CÔNG THỨC LƯƠNG"
   
   ✅ Hệ thống tự động:
      - Đọc JSON config
      - Map đến D:\Linh_Salary_Tool\2026\01\
      - Tìm tất cả file source
      - Nếu file tồn tại → Tạo formula linking
      - Nếu file chưa có → Để trống + Log cảnh báo
```

---

## 🔧 Cách Chỉnh Sửa Config

### Thêm Template Mới

1. Mở `Template\salary_sources.json`
2. Thêm entry mới trong `templates`:
   ```json
   "gt/new_template.xlsx": {
     "sheets": {
       "Bang TH nop": {
         "start_row": 9,
         "row_step": 2,
         "columns": {
           "D": {
             "source_file": "...",
             "source_sheet": "...",
             "source_range": "...",
             "column_index": 42
           }
         }
       }
     }
   }
   ```
3. Lưu file

### Thêm Cột Mới

1. Trong `"columns"` của sheet:
   ```json
   "E": {
     "source_file": "...",
     "source_sheet": "...",
     "source_range": "...",
     "column_index": 29
   }
   ```

### Thay Đổi File Source

1. Tìm cột cần chỉnh
2. Sửa `source_file` hoặc `source_sheet`
3. Lưu JSON
4. Rerun Tab 5

**Không cần sửa code Python!** 🎉

---

## 📊 Logs & Warnings

### Khi Chạy Tab 5, Nhật Ký Sẽ Hiển Thị:

```
📊 Bắt đầu triển khai công thức lương
📋 Template: gt/bep_an_va_cong_vu_template.xlsx
📂 Thư mục: D:\Scan2Excel-AI\01_danh_sach_chia_to\2026\01\gt
📅 Tháng/Năm: 01/2026
🔧 Đọc config từ: D:\Scan2Excel-AI\Template\salary_sources.json

🔍 Tìm thấy 2 file Excel để xử lý

📄 [1/2] Xử lý: Bếp ăn - 01.xlsx
   📋 Sheet: Bang TH nop
   👥 Nhân viên: 15 dòng (row 9)
      📌 Cột D: Công thức
      ✅ Cột E: Link từ Danh sách CBCNV.xlsx!Danh sách
      ⚠️ Cột G: File 'Chấm công 01.26.xlsx' không tìm thấy → Để trống
   📋 Sheet: Lương 01
   👥 Nhân viên: 15 dòng (row 10)
      📌 Cột H: Công thức
      ✅ Cột Y: Link từ CK tháng 12.xlsx!Bản gốc T12
   ✅ Đã lưu file

✅ Hoàn tất! Đã triển khai cho 2/2 file.

⚠️ Có 1 cảnh báo - Kiểm tra log ở trên.
```

---

## ⚙️ Cách Sửa Lỗi

### Lỗi: "Sheet không tìm thấy"
→ Kiểm tra tên sheet trong JSON khớp với file Excel

### Lỗi: "File không tìm thấy"
→ Kiểm tra đường dẫn D:\Linh_Salary_Tool\{YYYY}\{MM}\ có file tương ứng không

### Công thức linking không hoạt động
→ File source chưa tồn tại hoặc sheet tên sai

---

## 💡 Mẹo & Trik

### 1. Kiểm Tra Config Trước Khi Chạy
```python
from logic.salary_config_parser import get_salary_config
config = get_salary_config()
is_valid, errors = config.validate_template("gt/bep_an_va_cong_vu_template.xlsx")
if errors:
    print(f"Lỗi: {errors}")
```

### 2. Test File Source
```python
config = get_salary_config()
path = config.resolve_file_path("Danh sách CBCNV bản dùng làm lương T{mm_int}.xlsx", 2026, 1)
print(f"Found: {path}")  # Nếu None = file chưa có
```

### 3. Reuse Template Config
Nếu 2 template có cấu trúc giống → Có thể copy config từ template này sang template khác

---

## 🎯 Tổng Kết

| Điều | Trước | Sau |
|---|---|---|
| **Config** | Code Python (khó sửa) | JSON file (dễ sửa) |
| **Thêm template** | Sửa Python code | Thêm entry JSON |
| **Thay đổi file source** | Sửa code | Edit JSON + Save |
| **Handle missing files** | Treo/lỗi | Để trống + Log cảnh báo |
| **Số bước triển khai** | 3-4 bước | **1 lần nhấn nút** ✨ |

---

## 📝 Notes

- JSON config tự động load từ `Template\salary_sources.json`
- Hỗ trợ unlimited templates + sheets + columns
- Mỗi file Excel có thể có multiple sheets
- Mỗi sheet có thể có multiple columns với config khác nhau
- Formula & source file config đều hỗ trợ
