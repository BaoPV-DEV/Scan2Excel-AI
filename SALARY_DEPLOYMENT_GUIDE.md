# 📊 Hệ Thống Triển Khai Lương Tập Trung (JSON Config + Metadata) - Tab 5

## ✨ Tổng Quan (v2.1 - Metadata-Based)

**Cách tiếp cận mới & chính xác nhất:**
- **Tab 2** tách nhân viên → **Lưu metadata** (tab2_split_metadata.json) ghi lại file nào dùng template nào
- **Tab 5** đọc **metadata từ Tab 2** → Ánh xạ template tự động → Triển khai công thức

**Ưu điểm:**
✅ **Linh hoạt**: Mỗi tháng có các tổ khác nhau, metadata sẽ tự động cập nhật
✅ **Chính xác**: Không phụ thuộc vào hardcoded mapping
✅ **Tự động**: Tab 5 không cần chọn template, tất cả đã biết từ Tab 2

---

## 🔄 Workflow (v2.1)

## 🗂️ Cấu Trúc Thư Mục & Mapping

### **Nguồn Dữ Liệu (Input)**
```
D:\Linh_Salary_Tool\
├── 2026\
│   ├── 01\, 02\, ..., 12\ (Dữ liệu theo tháng)
│   │   ├── Danh sách CBCNV bản dùng làm lương T1.xlsx
│   │   ├── Danh sách CBCNV làm lương tháng 01.xlsx
│   │   ├── Chấm công 01.26.xlsx
│   │   ├── CK tháng 12.2025.xlsx (dữ liệu tháng trước)
│   │   ├── phụ cấp con nhỏ năm 2026.xlsx
│   │   └── ...
```

### **Output từ Tab 2 (Tách Nhân Viên)**
```
D:\Linh_Salary_Tool\01_danh_sach_chia_to\
├── 2026\
│   └── 04\
│       ├── gt\ (Admin/Security Teams)
│       │   ├── Bảo vệ - 04.xlsx              → gt/bao_ve_template.xlsx
│       │   ├── Bếp ăn - 04.xlsx             → gt/bep_an_va_cong_vu_template.xlsx
│       │   ├── Công vụ - 04.xlsx            → gt/bep_an_va_cong_vu_template.xlsx
│       │   ├── Bốc vác - 04.xlsx            → gt/boc_vac_template.xlsx
│       │   ├── C.điện - 04.xlsx             → gt/co_dien_template.xlsx
│       │   ├── K.hoạch - 04.xlsx            → gt/ke_hoach_template.xlsx
│       │   ├── K.thuật - 04.xlsx            → gt/ky_thuat_template.xlsx
│       │   ├── VP - 04.xlsx                 → gt/van_phong_template.xlsx
│       │   └── Kiểm hóa - 04.xlsx           → sx/kiem_hoa_template.xlsx
│       │
│       └── sx\ (Sewing/Production Teams)
│           ├── Tổ 1 - 04.xlsx               → sx/to_may_template.xlsx
│           ├── Tổ 2 - 04.xlsx               → sx/to_may_template.xlsx
│           ├── Tổ 3 - 04.xlsx               → sx/to_may_template.xlsx
│           ├── ...
│           ├── Tổ 12 - 04.xlsx              → sx/to_may_template.xlsx
│           ├── Tổ trưởng - 04.xlsx          → sx/to_may_template.xlsx
│           ├── Thời vụ - 04.xlsx            → sx/to_may_template.xlsx
│           ├── Cắt - 04.xlsx                → sx/to_may_template.xlsx
│           ├── H.thiện - 04.xlsx            → sx/to_may_template.xlsx
│           ├── Điều động - 04.xlsx          → sx/to_may_template.xlsx
│           ├── Hậu giặt - 04.xlsx           → sx/to_may_template.xlsx
│           ├── In - 04.xlsx                 → sx/to_may_template.xlsx
│           └── Kiểm hóa - 04.xlsx           → sx/kiem_hoa_template.xlsx
│
└── Template\ (Template Files với cấu hình)
    ├── salary_sources.json  ✨ **CONFIG TẬP TRUNG** (chỉnh sửa ở đây)
    ├── gt\
    │   ├── bao_ve_template.xlsx
    │   ├── bep_an_va_cong_vu_template.xlsx
    │   ├── boc_vac_template.xlsx
    │   ├── co_dien_template.xlsx
    │   ├── ke_hoach_template.xlsx
    │   ├── ky_thuat_template.xlsx
    │   └── van_phong_template.xlsx
    │
    └── sx\
        ├── to_may_template.xlsx
        └── kiem_hoa_template.xlsx
```

---

## 📝 JSON Config Structure

File: `Template\salary_sources.json` (Version 2.0)

```json
{
  "config_version": "2.0",
  "base_path": "D:\\Linh_Salary_Tool",
  "split_output_path": "D:\\Linh_Salary_Tool\\01_danh_sach_chia_to",
  "year_month_pattern": "{YYYY}\\{MM}",
  
  "template_mapping": {
    "gt/Bảo vệ": "gt/bao_ve_template.xlsx",
    "gt/Bếp ăn": "gt/bep_an_va_cong_vu_template.xlsx",
    "gt/Công vụ": "gt/bep_an_va_cong_vu_template.xlsx",
    "gt/Bốc vác": "gt/boc_vac_template.xlsx",
    "gt/C.điện": "gt/co_dien_template.xlsx",
    "gt/K.hoạch": "gt/ke_hoach_template.xlsx",
    "gt/K.thuật": "gt/ky_thuat_template.xlsx",
    "gt/VP": "gt/van_phong_template.xlsx",
    "gt/Kiểm hóa": "sx/kiem_hoa_template.xlsx",
    "sx/Tổ": "sx/to_may_template.xlsx",
    "sx/Cắt": "sx/to_may_template.xlsx",
    "sx/H.thiện": "sx/to_may_template.xlsx",
    "sx/Điều động": "sx/to_may_template.xlsx",
    "sx/Hậu giặt": "sx/to_may_template.xlsx",
    "sx/In": "sx/to_may_template.xlsx",
    "sx/Kiểm hóa": "sx/kiem_hoa_template.xlsx"
  },
  
  "templates": {
    "gt/bao_ve_template.xlsx": {
      "description": "Bảo vệ (Security)",
      "output_file_pattern": "gt/Bảo vệ - {MM}.xlsx",
      "sheets": {
        "Bang TH nop": { ... },
        "Lương {mm}": { ... }
      }
    },
    "sx/to_may_template.xlsx": {
      "description": "Tổ may (Sewing Teams)",
      "output_file_pattern": "sx/(Tổ|Tổ trưởng|Thời vụ|Cắt|H.thiện|...) .* - {MM}.xlsx",
      "sheets": { ... }
    }
    ...
  }
}
```

---

---

## 📊 Placeholder Hỗ Trợ

| Placeholder | Ví dụ | Ghi Chú |
|---|---|---|
| `{yyyy}` | 2026 | Năm hiện tại |
| `{yy_short}` | 26 | 2 chữ số cuối năm |
| `{mm}` | 04 | Tháng (2 chữ số) |
| `{mm_int}` | 4 | Tháng (số) |
| `{prev_year}` | 2025 | Năm trước |
| `{prev_month}` | 03 | Tháng trước (2 chữ số) |
| `{prev_month_int}` | 3 | Tháng trước (số) |

---

## 🎯 Chi Tiết Mapping Từng Tổ/Bộ Phận

### **GT (Admin/Security Teams) Templates**

| Tổ/Bộ Phận | File Output | Template | Nơi Lưu |
|---|---|---|---|
| Bảo vệ | `Bảo vệ - 04.xlsx` | `gt/bao_ve_template.xlsx` | `gt/` |
| Bếp ăn | `Bếp ăn - 04.xlsx` | `gt/bep_an_va_cong_vu_template.xlsx` | `gt/` |
| Công vụ | `Công vụ - 04.xlsx` | `gt/bep_an_va_cong_vu_template.xlsx` | `gt/` |
| Bốc vác | `Bốc vác - 04.xlsx` | `gt/boc_vac_template.xlsx` | `gt/` |
| Cơ điện | `C.điện - 04.xlsx` | `gt/co_dien_template.xlsx` | `gt/` |
| Kế hoạch | `K.hoạch - 04.xlsx` | `gt/ke_hoach_template.xlsx` | `gt/` |
| Kỹ thuật | `K.thuật - 04.xlsx` | `gt/ky_thuat_template.xlsx` | `gt/` |
| Văn phòng | `VP - 04.xlsx` | `gt/van_phong_template.xlsx` | `gt/` |
| Kiểm hóa | `Kiểm hóa - 04.xlsx` | `sx/kiem_hoa_template.xlsx` | `sx/` |

### **SX (Sewing/Production Teams) Templates**

| Tổ/Bộ Phận | File Output | Template | Nơi Lưu |
|---|---|---|---|
| Tổ 1-12 | `Tổ 1 - 04.xlsx`, `Tổ 2 - 04.xlsx`, ... | `sx/to_may_template.xlsx` | `sx/` |
| Tổ trưởng | `Tổ trưởng - 04.xlsx` | `sx/to_may_template.xlsx` | `sx/` |
| Thời vụ | `Thời vụ - 04.xlsx` | `sx/to_may_template.xlsx` | `sx/` |
| Cắt | `Cắt - 04.xlsx` | `sx/to_may_template.xlsx` | `sx/` |
| Hoàn thiện | `H.thiện - 04.xlsx` | `sx/to_may_template.xlsx` | `sx/` |
| Điều động | `Điều động - 04.xlsx` | `sx/to_may_template.xlsx` | `sx/` |
| Hậu giặt | `Hậu giặt - 04.xlsx` | `sx/to_may_template.xlsx` | `sx/` |
| In | `In - 04.xlsx` | `sx/to_may_template.xlsx` | `sx/` |
| Kiểm hóa | `Kiểm hóa - 04.xlsx` | `sx/kiem_hoa_template.xlsx` | `sx/` |

---

## 🔍 Cách Tab 5 Hoạt Động

### **Step 1: Quét Output từ Tab 2**
```python
# Tab 5 quét thư mục: 
# D:\Linh_Salary_Tool\01_danh_sach_chia_to\2026\04\

# Tìm được các file:
# - gt/Bảo vệ - 04.xlsx
# - gt/Bếp ăn - 04.xlsx
# - sx/Tổ 1 - 04.xlsx
# - sx/Tổ 2 - 04.xlsx
# - ...
```

### **Step 2: Xác Định Template**
```python
# Dựa trên tên folder + tên file, ánh xạ với template:

# gt/Bảo vệ - 04.xlsx
#   ↓ Lookup trong template_mapping
#   ↓ Match: "gt/Bảo vệ" → "gt/bao_ve_template.xlsx"

# sx/Tổ 1 - 04.xlsx
#   ↓ Match: "sx/Tổ" → "sx/to_may_template.xlsx"
```

### **Step 3: Load Config từ JSON**
```python
# Tìm template "gt/bao_ve_template.xlsx" trong salary_sources.json
# Load các cấu hình:
#   - Sheets cần xử lý: "Bang TH nop", "Lương 04"
#   - Rows: start_row = 9, row_step = 2
#   - Columns: D, E, F, ... và source files tương ứng
```

### **Step 4: Triển Khai Công Thức**
```python
# Ánh xạ dữ liệu từ source files:
# 1. Đọc: D:\Linh_Salary_Tool\2026\04\Danh sách CBCNV...xlsx
# 2. Ghi vào: D:\Linh_Salary_Tool\01_danh_sach_chia_to\2026\04\gt\Bảo vệ - 04.xlsx
# 3. Tích hợp: CK tháng 03.2026.xlsx, phụ cấp con nhỏ, v.v.
```

---

## ⚙️ Cấu Hình Thêm & Chỉnh Sửa

### 1️⃣ Thêm Template Mới (nếu cần)

1. Đặt file template mới vào `Template/gt/` hoặc `Template/sx/`
2. Mở `Template\salary_sources.json`
3. Thêm mapping trong `template_mapping`:
   ```json
   "gt/New Team": "gt/new_template.xlsx"
   ```
4. Thêm cấu hình chi tiết trong `templates`:
   ```json
   "gt/new_template.xlsx": {
     "description": "New Team Description",
     "output_file_pattern": "gt/New Team - {MM}.xlsx",
     "sheets": {
       "Bang TH nop": {
         "start_row": 9,
         "row_step": 2,
         "columns": {
           "D": {
             "source_file": "...",
             "source_sheet": "...",
             "source_range": "$D$1:$AS$1000",
             "column_index": 42
           }
         }
       }
     }
   }
   ```

### 2️⃣ Thêm Cột Dữ Liệu Mới

Trong `"columns"` của mỗi sheet:
```json
"E": {
  "source_file": "Danh sách CBCNV làm lương tháng {mm}.xlsx",
  "source_sheet": "Danh Sách dùng",
  "source_range": "$E$5:$AG$11000",
  "column_index": 29,
  "description": "Ngày vào"
}
```

### 3️⃣ Thêm Formula (Công Thức Tính)

Nếu cần tính toán thay vì ánh xạ:
```json
"H": {
  "formula": "=+(F{row}/26/8*{bep_cong_vu_rate})*H{row_minus_1}",
  "type": "formula",
  "description": "Tính lương thêm giờ"
}
```

Variables trong formula:
- `{row}` - Số dòng hiện tại
- `{row_minus_1}` - Dòng trước (row - 1)
- `{bep_cong_vu_rate}` - Tỷ lệ đặc biệt cho Bếp ăn

---

## ✅ Checklist Chuẩn Bị Tab 5

Trước khi nhấn "TRIỂN KHAI CÔNG THỨC":

- [ ] ✅ **Tab 2 hoàn thành**: File split đã lưu tại `D:\Linh_Salary_Tool\01_danh_sach_chia_to\2026\04\`
- [ ] ✅ **Dữ liệu Source sẵn sàng** trong `D:\Linh_Salary_Tool\2026\04\`:
  - Danh sách CBCNV bản dùng làm lương T4.xlsx
  - Danh sách CBCNV làm lương tháng 04.xlsx
  - Chấm công 04.26.xlsx
  - CK tháng 03.2026.xlsx (tháng trước)
  - phụ cấp con nhỏ năm 2026.xlsx
- [ ] ✅ **Template files** có sẵn trong `Template/gt/` và `Template/sx/`
- [ ] ✅ **salary_sources.json** đã cấu hình đầy đủ (v2.0)
- [ ] ✅ **Tháng/Năm** được chọn chính xác (04/2026)

---

## 🚀 Quy Trình Triển Khai (Step-by-Step)

```
┏━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━┓
┃ STEP 1: Chuẩn Bị Dữ Liệu Source                    ┃
┗━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━┛
D:\Linh_Salary_Tool\2026\04\ phải có:
✓ Danh sách CBCNV bản dùng làm lương T4.xlsx
✓ Chấm công 04.26.xlsx
✓ CK tháng 03.2026.xlsx
✓ phụ cấp con nhỏ năm 2026.xlsx
                  ↓
┏━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━┓
┃ STEP 2: Chạy Tab 2 - Tách Nhân Viên                ┃
┗━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━┛
Output: D:\Linh_Salary_Tool\01_danh_sach_chia_to\2026\04\
  ├── gt/ (8 teams)
  └── sx/ (18 teams + variants)
                  ↓
┏━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━┓
┃ STEP 3: Tab 5 - Triển Khai Lương (ONE-CLICK)       ┃
┗━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━┛
1. Chọn thư mục: 01_danh_sach_chia_to\2026\04\
2. Chọn tháng: 04
3. Chọn năm: 2026
4. Nhấn "TRIỂN KHAI CÔNG THỨC LƯƠNG"

Hệ thống tự động:
  ✓ Quét tất cả file trong gt/ và sx/
  ✓ Ánh xạ với template (dựa trên tên file)
  ✓ Load config từ salary_sources.json
  ✓ Tích hợp dữ liệu từ D:\Linh_Salary_Tool\2026\04\
  ✓ Tạo formula linking
  ✓ Log tất cả thao tác + cảnh báo nếu có file missing
                  ↓
                ✅ HOÀN THÀNH
```

---

## 📱 Lỗi & Cách Khắc Phục

| Lỗi | Nguyên Nhân | Cách Khắc Phục |
|---|---|---|
| File source không tìm thấy | File chưa có tại `D:\Linh_Salary_Tool\2026\04\` | Chuẩn bị file source & chạy lại |
| Template mapping không match | Tên file output không match template_mapping | Kiểm tra tên file, sửa salary_sources.json |
| Formula không tạo được | Column không định nghĩa trong JSON | Thêm cấu hình cho column đó trong JSON |
| Dữ liệu trống | source_range sai | Kiểm tra source_sheet + column_index |

---

## 📞 Hỗ Trợ

- ❓ Kiểm tra log output trong **Tab 5 UI**
- 🔍 Xem chi tiết cấu hình tại `Template\salary_sources.json`
- 📋 So sánh tên file & thư mục với `template_mapping`
- 💾 Backup các file trước khi chạy

---

## 📝 Version History

| Version | Ngày | Thay Đổi |
|---|---|---|
| v2.0 | 2026-05-23 | ✨ Thêm template_mapping, split_output_path, đầy đủ GT templates + SX templates |
| v1.0 | 2026-05 | Khởi tạo hệ thống JSON-based deployment |

,
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
