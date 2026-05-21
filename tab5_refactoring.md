# 📋 Tài Liệu Refactoring - Tách Tab 1 và Tab 5

## 📌 Tổng Quan Thay Đổi

Để tối ưu hóa cấu trúc ứng dụng, chúng tôi đã thực hiện tách biệt chức năng:
- **Tab 1 (Tách File Nhân Viên)**: Tập trung vào phân chia danh sách nhân viên theo tổ - KHÔNG CÓ linking ngoài
- **Tab 5 (Triển Khai Lương)** [MỚI]: Xử lý linking công thức lương với dữ liệu ngoài

---

## 🔄 Quy Trình Công Việc

### Quy trình cũ (Tab 1 xử lý toàn bộ):
```
Tab 1: Tách File NV
   ├── Phân loại nhân viên
   ├── Copy template
   ├── Ghi dữ liệu
   ├── Lưu file
   └── ❌ Áp dụng công thức lương (linking) - RỦI RO: file chưa tồn tại
```

### Quy trình mới (Tab 1 + Tab 5):
```
Tab 1: Tách File NV (Đơn giản hóa)
   ├── Phân loại nhân viên
   ├── Copy template
   ├── Ghi dữ liệu
   └── Lưu file

    ↓ (File đã tạo xong, sẵn sàng)

Tab 5: Triển Khai Lương (Riêng biệt)
   ├── Chọn thư mục chứa file Excel
   ├── Chọn template để áp dụng linking
   ├── Chọn tháng/năm
   └── Áp dụng công thức lương với linking ngoài
       (Chỉ khi tất cả dữ liệu ngoài đã sẵn sàng)
```

---

## 📁 Cấu Trúc Thư Mục Mới

```
ui/
├── tab1_split/              # Xử lý tách file nhân viên
│   ├── __init__.py
│   ├── split_widget.py      # Giao diện (UI)
│   └── split_thread.py      # Xử lý ngầm (threading)
│
├── tab5_salary_deploy/      # ✨ MỚI: Triển khai lương
│   ├── __init__.py
│   ├── salary_widget.py     # Giao diện (UI)
│   └── salary_thread.py     # Xử lý ngầm (threading)
│
└── ... (các tab khác)

logic/
├── tab1_writer.py           # ✏️ Đã sửa: Bỏ TEMPLATE_FORMULA_CONFIG
├── tab1_classifier.py       # (không thay đổi)
├── salary_config.py         # ✨ MỚI: Config công thức lương
└── salary_deployer.py       # ✨ MỚI: Logic triển khai lương
```

---

## 🔧 Chi Tiết Thay Đổi

### 1. Tab 1: Đơn giản hóa (Loại bỏ linking)

**File: `logic/tab1_writer.py`**
- ❌ Xóa: `TEMPLATE_FORMULA_CONFIG` (6000+ dòng config công thức lương)
- ❌ Xóa: `apply_custom_formulas()` (hàm áp dụng công thức)
- ❌ Xóa: `apply_custom_formulas_openpyxl()` (hàm dự phòng openpyxl)
- ✏️ Sửa: `write_all_groups()` - Bỏ gọi hàm áp dụng công thức

**Lợi ích:**
- Tab 1 chạy nhanh hơn (không chặn chờ external files)
- Ít lỗi hơn (không cần kiểm tra file ngoài tồn tại)
- Dễ bảo trì hơn (logic đơn giản hơn)

---

### 2. Tab 5: Triển khai lương (MỚI)

**Files:**
- `logic/salary_config.py` - Lưu toàn bộ config công thức lương
- `logic/salary_deployer.py` - Logic áp dụng công thức
- `ui/tab5_salary_deploy/salary_widget.py` - Giao diện
- `ui/tab5_salary_deploy/salary_thread.py` - Xử lý ngầm

**Tính năng:**
- ✅ Chọn thư mục chứa file Excel (từ Tab 1 hoặc bất kỳ đâu)
- ✅ Chọn template (Bảo vệ, Bếp ăn, Công vụ, v.v.)
- ✅ Chọn tháng/năm
- ✅ Áp dụng công thức lương tự động
- ✅ Hỗ trợ multiple files cùng lúc
- ✅ Nhật ký chi tiết

---

## 📊 Công Thức Lương Được Hỗ Trợ

### Template: `gt/bao_ve_template.xlsx`
- Sheet "Bang TH nop": Lookup dữ liệu từ "Danh sách CBCNV bản dùng làm lương"
- Sheet "Lương {mm}": Lookup từ file "CK tháng trước"

### Template: `gt/bep_an_va_cong_vu_template.xlsx`
- 10+ cột với linking đến:
  - Danh sách CBCNV
  - Chấm công
  - Phụ cấp con nhỏ
  - Thâm niên
  - v.v.

---

## 🎯 Hướng Sử Dụng

### Quy trình 1 tháng (từ file gốc đến file lương cuối cùng):

```
Step 1: Chuẩn bị dữ liệu ngoài
   - Danh sách CBCNV làm lương
   - File chấm công
   - Phụ cấp con nhỏ
   - v.v. (tất cả ở sẵn)

Step 2: Tab 1 - Tách File Nhân Viên
   - Chọn file danh sách NV gốc
   - Chọn tháng/năm
   - Click "THỰC HIỆN TÁCH FILE"
   - ✅ File split ra theo tổ

Step 3: Tab 5 - Triển Khai Công Lương
   - Chọn thư mục chứa file split
   - Chọn template phù hợp
   - Chọn tháng/năm
   - Click "TRIỂN KHAI CÔNG THỨC LƯƠNG"
   - ✅ File sẵn sàng với công thức linking
```

---

## ⚠️ Lưu Ý Quan Trọng

1. **Thứ tự xử lý**: 
   - Tab 1 phải chạy TRƯỚC Tab 5
   - Tab 5 phải chạy SAU khi tất cả dữ liệu ngoài sẵn sàng

2. **Đường dẫn dữ liệu ngoài**:
   - Cần đảm bảo đường dẫn trong `salary_config.py` khớp với hệ thống thực tế
   - Ví dụ: `D:\HỒ SƠ+LƯƠNG SD\LƯƠNG\...`

3. **Phiên bản Excel**:
   - Sử dụng Excel 2016 trở lên
   - Hỗ trợ .xlsx (không hỗ trợ .xls cũ)

---

## 🔍 Xử Lý Sự Cố

### Tab 5 báo lỗi "Template không được hỗ trợ"
→ Kiểm tra template_key trong `salary_config.py` có đúng format không

### Tab 5 báo "Không tìm thấy sheet"
→ Kiểm tra tên sheet trong file Excel khớp với config

### Công thức linking không hoạt động sau triển khai
→ Kiểm tra dữ liệu ngoài (Danh sách CBCNV, Chấm công) có đúng tên sheet và vị trí ô không

---

## 📈 Lợi Ích của Refactoring

| Tiêu chí | Trước | Sau |
|---------|-------|-----|
| **Tốc độ Tab 1** | Chậm (phụ thuộc linking) | Nhanh (không linking) |
| **Tính linh hoạt** | Cứng nhắc | Linh hoạt (có thể áp dụng linking sau) |
| **Bảo trì code** | Phức tạp | Đơn giản (tách biệt) |
| **Debug lỗi** | Khó (nhiều yếu tố) | Dễ (từng tab riêng) |
| **Reuse code** | Khó | Dễ (Tab 5 có thể dùng cho bất kỳ file nào) |

---

## 🚀 Tương Lai

Những cải tiến có thể trong tương lai:
- [ ] Hỗ trợ thêm template khác
- [ ] Lưu preset (combo template + tháng/năm)
- [ ] Batch processing (cùng lúc nhiều tháng)
- [ ] Export report linking status
- [ ] Validation công thức trước khi áp dụng
