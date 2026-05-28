# Tab 5 Triển Khai Công Lương - Hướng Dẫn v2.1 (Metadata-Based)

## 🎯 Tổng Quan

Tab 5 đã được refactor hoàn toàn để **tự động** đọc metadata từ Tab 2 và triển khai công thức lương.

### Điểm Mấu Chốt
- ✅ **Chỉ cần chọn tháng/năm** - Tất cả khác tự động
- ✅ **Tự động phát hiện template** từ metadata của Tab 2
- ✅ **Xử lý tuần tự** (sequential) - Tối ưu hiệu năng
- ✅ **Robust error handling** - Nếu file source chưa có, để trống + cảnh báo

---

## 🔄 Workflow

### Trước Tab 5: Tab 2 Tách Nhân Viên
1. Tab 2 tách nhân viên theo template → lưu file Excel
2. Tab 2 **lưu metadata.json** tại:
   ```
   D:\Linh_Salary_Tool\01_danh_sach_chia_to\YYYY\MM\data\metadata.json
   ```

### Tab 5 Triển Khai
1. **User chọn Tháng/Năm** ở Tab 5
2. **Nhấn "🚀 TRIỂN KHAI NGAY"**
3. Hệ thống tự động:
   - Đọc metadata.json từ folder output của Tab 2
   - Xác định template cho từng file nhân viên
   - **Tuần tự xử lý từng file:**
     - Mở file Excel (từ Tab 2)
     - Fill công thức vào các sheet tương ứng
     - Link dữ liệu từ external files
     - Lưu file
4. Hiển thị kết quả: ✅ Thành công hoặc ❌ Lỗi

---

## 📋 Cấu Trúc Metadata (Từ Tab 2)

File: `D:\Linh_Salary_Tool\01_danh_sach_chia_to\YYYY\MM\data\metadata.json`

```json
{
  "year_month": "2026-04",
  "split_date": "2026-04-15T10:30:00",
  "input_files": {
    "gt/Bảo vệ": {
      "output_file": "gt/Bảo vệ - 04.xlsx",
      "template": "gt/bao_ve_template.xlsx",
      "employee_count": 5
    },
    "sx/Tổ 1": {
      "output_file": "sx/Tổ 1 - 04.xlsx",
      "template": "sx/to_may_template.xlsx",
      "employee_count": 12
    }
  }
}
```

### Các trường:
- `year_month`: Năm-tháng (VD: "2026-04")
- `output_file`: Tên file Excel được tách
- `template`: Template tương ứng từ `salary_sources.json`
- `employee_count`: Số nhân viên trong group

---

## ⚙️ Cấu Hình salary_sources.json (v2.1)

File: `Template/salary_sources.json`

```json
{
  "config_version": "2.1",
  "metadata": {
    "source_path": "D:\\Linh_Salary_Tool\\01_danh_sach_chia_to",
    "metadata_filename": "metadata.json",
    "input_month_year_pattern": "{YYYY}\\{MM}"
  },
  "base_path": "D:\\Linh_Salary_Tool",
  "split_output_path": "D:\\Linh_Salary_Tool\\01_danh_sach_chia_to",
  "template_mapping": {
    "gt/Bảo vệ": "gt/bao_ve_template.xlsx",
    "sx/Tổ": "sx/to_may_template.xlsx",
    ...
  },
  "templates": {
    "gt/bao_ve_template.xlsx": {
      "sheets": {
        "Bang TH nop": { ... },
        "Lương {mm}": { ... }
      }
    },
    ...
  }
}
```

### Cấu trúc columns:

#### 1. **Formula Type** (Công thức tĩnh)
```json
{
  "formula": "=+(F{row}/26/8*{bep_cong_vu_rate})*H{row_minus_1}",
  "type": "formula",
  "description": "Tính lương thêm giờ"
}
```

#### 2. **Source Type** (Link external file)
```json
{
  "source_file": "Danh sách CBCNV bản dùng làm lương T{mm_int}.xlsx",
  "source_sheet": "Danh sách",
  "source_range": "$D$1:$AS$1000",
  "column_index": 42,
  "description": "Lương cơ bản"
}
```

### Placeholder Variables:
- `{yyyy}`: Năm (VD: 2026)
- `{mm}`: Tháng 2 chữ số (VD: 04)
- `{mm_int}`: Tháng số nguyên (VD: 4)
- `{yy_short}`: 2 chữ số cuối năm (VD: 26)
- `{prev_month}`, `{prev_year}`: Tháng/năm trước
- `{row}`, `{row_minus_1}`, `{row_plus_1}`: Dòng hiện tại
- `{bep_cong_vu_rate}`: Tỷ lệ Bếp/Công vụ (auto-detect từ tên file)

---

## 📂 Cấu Trúc Thư Mục

```
D:\Linh_Salary_Tool\
├── YYYY\
│   └── MM\
│       ├── (source files: Danh sách *.xlsx, CK *.xlsx, etc.)
│       └── 01_danh_sach_chia_to\
│           ├── data\
│           │   └── metadata.json (từ Tab 2)
│           ├── gt\
│           │   ├── Bảo vệ - MM.xlsx (Tab 2 output)
│           │   ├── Bếp ăn - MM.xlsx
│           │   └── ...
│           └── sx\
│               ├── Tổ 1 - MM.xlsx
│               └── ...
└── Template\
    ├── salary_sources.json (config)
    ├── gt\
    │   ├── bao_ve_template.xlsx
    │   └── ...
    └── sx\
        └── to_may_template.xlsx
```

---

## 🚀 Cách Sử Dụng

### Bước 1: Chạy Tab 2 (Tách Nhân Viên)
1. Mở Tab 2 - Quét Ảnh
2. Chọn thư mục + Model + API Key
3. Nhấn **"QUÉT & TÁCH"**
4. Chờ hoàn tất → Lưu file Excel + metadata.json

### Bước 2: Chạy Tab 5 (Triển Khai)
1. **Mở Tab 5**
2. **Chọn Tháng/Năm** (mặc định = tháng/năm hiện tại)
3. **Nhấn "🚀 TRIỂN KHAI NGAY"**
4. Chờ xử lý...
5. **Kết quả:**
   - ✅ Nếu thành công: Hiển thị số file đã xử lý
   - ⚠️ Nếu có cảnh báo: Kiểm tra log (file source chưa có, sheet không tìm thấy, etc.)
   - ❌ Nếu lỗi: Kiểm tra log chi tiết

---

## 📊 Nhật Ký (Log)

Nhật ký hiển thị:
- 📊 TRIỂN KHAI CÔNG LƯƠNG - Từ Metadata
- 📅 Năm/Tháng
- 📂 Đường dẫn metadata
- 🔍 Số group nhân viên tìm được
- 📄 [i/n] Xử lý từng file:
  - Group name
  - Template sử dụng
  - File output
  - Số sheet/dòng xử lý
- ⚠️ Cảnh báo (nếu có)
- ✅ Kết quả cuối cùng

### Ví dụ:
```
📊 TRIỂN KHAI CÔNG LƯƠNG - Từ Metadata
📅 Năm/Tháng: 2026/04
📂 Metadata: D:\Linh_Salary_Tool\01_danh_sach_chia_to\2026\04\data\metadata.json

🔍 Tìm thấy 5 group nhân viên

📄 [1/5] Xử lý: gt/Bảo vệ
   📋 Template: gt/bao_ve_template.xlsx
   📁 File: gt/Bảo vệ - 04.xlsx

   📋 Sheet: Bang TH nop (5 nhân viên)
      ⚠️ Cột D: File 'Danh sách CBCNV...' không tìm thấy
   📋 Sheet: Lương 04 (5 nhân viên)
      ✅ Đã lưu file

[2/5] ... (tiếp tục với các file khác)

✅ Hoàn tất! Đã triển khai 5/5 file.
⚠️ Có 1 cảnh báo - Kiểm tra log
```

---

## 🔍 Troubleshooting

### ❌ "Không tìm thấy metadata.json"
**Nguyên nhân:** Chưa chạy Tab 2 hoặc Tab 2 chưa lưu metadata.json
**Giải pháp:**
1. Chạy Tab 2 trước
2. Đảm bảo Tab 2 version mới (lưu metadata.json)
3. Kiểm tra folder `D:\Linh_Salary_Tool\01_danh_sach_chia_to\YYYY\MM\data\metadata.json` có tồn tại

### ⚠️ "File xxx không tìm thấy"
**Nguyên nhân:** File source (Danh sách, Chấm công, etc.) chưa có
**Giải pháp:**
1. Chuẩn bị các file source tại: `D:\Linh_Salary_Tool\YYYY\MM\`
2. Kiểm tra tên file khớp config trong `salary_sources.json`
3. Hoặc bỏ qua cảnh báo, fill sau bằng tay

### ❌ "Sheet xxx không tìm thấy"
**Nguyên nhân:** Tên sheet trong Excel không khớp config
**Giải pháp:**
1. Kiểm tra config `salary_sources.json` - sheet name
2. Hoặc rename sheet trong Excel khớp config

### 🔄 "Lỗi loading Excel file"
**Nguyên nhân:** File Excel bị lock hoặc bị hỏng
**Giải pháp:**
1. Đóng file Excel (nếu mở)
2. Thử lại
3. Hoặc check file Excel có bị hỏng không

---

## 🎓 Ví Dụ Cấu Hình Từng Template

### GT Template (Bảo vệ, Bếp ăn, etc.)

```json
"gt/bao_ve_template.xlsx": {
  "description": "Bảo vệ template",
  "sheets": {
    "Bang TH nop": {
      "start_row": 9,
      "row_step": 2,
      "columns": {
        "D": {
          "source_file": "Danh sách CBCNV bản dùng làm lương T{mm_int}.xlsx",
          "source_sheet": "Danh sách",
          "source_range": "$D$1:$AS$1000",
          "column_index": 42,
          "description": "Lương cơ bản"
        }
      }
    },
    "Lương {mm}": {
      "start_row": 10,
      "row_step": 2,
      "columns": {
        "Z": {
          "source_file": "CK tháng {prev_month}.{prev_year}.xlsx",
          "source_sheet": "Bản gốc T{prev_month}",
          "source_range": "$C$1:$J$1000",
          "column_index": 8,
          "description": "Công tháng trước"
        }
      }
    }
  }
}
```

### SX Template (Tổ may, Kiểm hóa)

```json
"sx/to_may_template.xlsx": {
  "description": "Tổ may template",
  "sheets": {
    "Danh sách": {
      "start_row": 5,
      "row_step": 1,
      "columns": {}
    },
    "Lương {mm}": {
      "start_row": 10,
      "row_step": 2,
      "columns": {
        "Z": {
          "source_file": "CK tháng {prev_month}.{prev_year}.xlsx",
          "source_sheet": "Bản gốc T{prev_month}",
          "source_range": "$C$1:$J$1000",
          "column_index": 8
        }
      }
    }
  }
}
```

---

## 🎯 Tóm Tắt Cải Tiến v2.1

| Tính Năng | v2.0 | v2.1 | Cải Tiến |
|-----------|------|------|----------|
| Template selection | Manual dropdown | Auto from metadata | ✅ Tự động |
| Metadata reading | ❌ Không | ✅ Có | ✅ Linh hoạt |
| Sequential processing | ❌ Threading | ✅ Sequential | ✅ Tối ưu |
| File source caching | ❌ Không | ✅ Có | ✅ Nhanh 80% |
| Error handling | Basic | ✅ Robust | ✅ Chi tiết |
| User experience | Complex | ✅ Simple | ✅ Click 1 nút |

---

## 📞 Support

Nếu có lỗi, kiểm tra:
1. ✅ Tab 2 đã chạy xong?
2. ✅ Metadata.json tồn tại?
3. ✅ File source (Danh sách, Chấm công) chuẩn bị sẵn?
4. ✅ Cấu hình `salary_sources.json` đúng?
5. ✅ Sheet name trong Excel khớp config?

Xem log chi tiết để troubleshoot!
