# Tab 5: Triển Khai Công Thức Lương Vào Excel

## 📋 Mô Tả Chức Năng

Tab này là bước cuối trong quy trình. Nó triển khai **công thức lương** vào các file Excel đã được tích hợp (từ Tab 4). Công thức được định cấu hình từ template và tự động điền vào từng hàng nhân viên. Hỗ trợ công thức nội bộ, công thức tham chiếu từ file khác, và công thức tùy chỉnh theo tổ.

---

## 🎯 Quy Trình Làm Việc

### 1. **Chế Độ Chạy**

#### Chế Độ 1: Chạy Tất Cả Tổ
- **Radio button**: "Chạy tất cả các tổ"
- Ứng dụng sẽ quét **metadata.json** để lấy danh sách tất cả file cần xử lý
- Tự động xác định mỗi file dùng template nào
- Xử lý tất cả file Excel trong folder `01_danh_sach_chia_to\YYYY\MM\`

#### Chế Độ 2: Chạy 1 Tổ
- **Radio button**: "Chạy 1 tổ"
- Nhấn **"📂 Chọn Folder"** để chỉ định folder của 1 tổ cụ thể
- Ứng dụng sẽ **chỉ xử lý tổ được chọn**
- Tự động trích xuất năm/tháng từ đường dẫn folder

### 2. **Cấu Hình Thời Gian & Năm (Nếu Chế Độ Tất Cả)**
- Chọn **Tháng** (01 → 12)
- Chọn **Năm** (thường là năm hiện tại)
- Phải khớp với **Năm/Tháng ở Tab 2, 3, 4**

### 3. **Cấu Hình File Nguồn (Tuỳ Chọn)**
- Có 4 card cấu hình cho các file nguồn:
  - **Danh sách CBCNV Lương**: Chứa thông tin cơ bản nhân viên (dùng cho công thức liên kết)
  - **Chuyển Khoản Tháng Trước**: Dữ liệu lương tháng trước (nếu cần)
  - **Danh sách CBCNV Làm Lương Tháng**: Danh sách nhân viên làm lương tháng này
  - **Chấm Công**: Dữ liệu chấm công (số ngày công)
  - **Phụ Cấp Con Nhỏ**: Phụ cấp con nhỏ (nếu áp dụng)

- **Cách cấu hình**: Nhấn **"📂 Chọn File"** trên mỗi card để chỉ định file nguồn
- **Mục đích**: Các công thức VLOOKUP sẽ tham chiếu đến các file này

### 4. **Bắt Đầu Triển Khai**
- Nhấn nút **"🚀 BẮT ĐẦU TRIỂN KHAI"** (nút xanh)
- Ứng dụng sẽ:
  - **Đọc metadata.json** để xác định danh sách file & template
  - **Mở Excel COM** (hidden) với performance cao
  - **Mở sẵn file nguồn** để công thức tham chiếu không bị treo
  - **Với từng file Excel tổ**:
    - Mở file
    - Tìm sheet phù hợp (VD: "Bang TH nop")
    - Triển khai công thức vào từng cột theo cấu hình template
    - Lưu file
  - **Ghi log chi tiết** cho từng bước

### 5. **Theo Dõi Tiến Độ**
- Thanh tiến độ hiển thị % hoàn tất
- Logs chi tiết hiển thị:
  - File đang xử lý (số thứ tự / tổng)
  - Template sử dụng
  - Số sheet được triển khai công thức
  - Số hàng nhân viên được cập nhật
  - Cảnh báo hoặc lỗi (nếu có)

### 6. **Hoàn Tất**
- Thông báo **"✅ Hoàn tất! Đã triển khai {X}/{Y} file"**
- File Excel được cập nhật tự động
- **Nhấn F9 hoặc Ctrl+Alt+F9 trong Excel** để tính kết quả công thức

---

## 📁 File Sử Dụng

| Loại | File | Vị Trí |
|------|------|--------|
| **Input Metadata** | metadata.json | `D:\Linh_Salary_Tool\04_data\YYYY\MM\` |
| **Input Excel** | File Excel tách | `D:\Linh_Salary_Tool\01_danh_sach_chia_to\YYYY\MM\` |
| **File Nguồn** | Chấm công, Phụ cấp... | Bất kỳ (do user cấu hình) |
| **Output Excel** | File Excel với công thức | Cùng vị trí (01_danh_sach_chia_to) |
| **Config Công Thức** | salary_config.json | `Template/data/` |

---

## 🔧 Cấu Trúc Công Thức Template

### Ví Dụ: Công Thức Tính Lương Tổ Máy

```json
{
  "sheets": {
    "Bang TH nop": {
      "start_row": 9,
      "row_step": 2,
      "columns": {
        "D": {
          "type": "team_formula",
          "team_configs": {
            "Tổ Cắt|Cut": {
              "formula": "=C{row}*1.25",
              "description": "Lương Tổ Cắt (125% hệ số)"
            },
            "Tổ May|Sew": {
              "formula": "=C{row}*1.15",
              "description": "Lương Tổ May (115% hệ số)"
            }
          }
        },
        "E": {
          "type": "source",
          "file": "cham_cong_{yyyy}_{mm}.xlsx",
          "sheet": "Chấm Công",
          "range": "B:C",
          "column_index": 2
        }
      }
    }
  }
}
```

### 3 Loại Công Thức

#### 1. **Formula (Công Thức Nội Bộ)**
- Công thức tính toán chỉ dùng dữ liệu trong file Excel hiện tại
- VD: `=C{row}*1.25` (Số lượng × 1.25)

#### 2. **Source (Công Thức Tham Chiếu)**
- Sử dụng VLOOKUP để tham chiếu dữ liệu từ file khác
- VD: Lấy số ngày công từ file chấm công
- Công thức: `=IFERROR(VLOOKUP(C{row},[file]Sheet,col,0),"")`

#### 3. **Team Formula (Công Thức Theo Tổ)**
- Công thức khác nhau cho mỗi tổ
- Ứng dụng tự động phát hiện tên tổ từ tên file
- VD: Tổ Cắt dùng hệ số 1.25, Tổ May dùng 1.15

---

## 📊 Cấu Trúc Metadata.json

```json
{
  "year_month": "2026-04",
  "output_dir": "D:\\Linh_Salary_Tool\\01_danh_sach_chia_to\\2026\\04",
  "mapping": {
    "sx/Tổ Cắt - 04.xlsx": {
      "template": "sx/to_may_template.xlsx",
      "group_name": "Tổ Cắt",
      "employee_count": 50,
      "folder": "sx"
    },
    "gt/Bảo vệ - 04.xlsx": {
      "template": "gt/kiem_hoa_template.xlsx",
      "group_name": "Bảo vệ",
      "employee_count": 10,
      "folder": "gt"
    }
  }
}
```

---

## ⚠️ Lưu Ý Quan Trọng

### Metadata & Cấu Hình
- **metadata.json bắt buộc**: Phải chạy Tab 2 trước để tạo metadata
- **Vị trí metadata**: `D:\Linh_Salary_Tool\04_data\YYYY\MM\metadata.json`
- **salary_config.json**: Định nghĩa công thức cho từng template

### File Nguồn
- **Tuỳ chọn**: Nếu không cấu hình file nguồn, công thức VLOOKUP sẽ bị bỏ qua
- **Đúng format**: File nguồn phải có sheet và range đúng như cấu hình
- **Mã nhân viên khớp**: Mã nhân viên trong file Excel phải khớp với file nguồn (nếu dùng VLOOKUP)

### Performance
- **Excel COM mở ẩn**: Ứng dụng mở Excel ẩn (không hiển thị lên màn hình)
- **Tối ưu hiệu suất**: Mở sẵn file nguồn, tắt tính toán tự động, tắt cập nhật màn hình
- **Restart định kỳ**: Mỗi 5 file sẽ restart Excel để tránh memory leak

### Sau Hoàn Tất
- **Tính kết quả**: Phải **nhấn F9 hoặc Ctrl+Alt+F9 trong Excel** để tính công thức
- **Kiểm tra dữ liệu**: Xem file Excel sau lần chạy đầu tiên để đảm bảo công thức chính xác
- **Không mở file**: Khi chạy Tab 5, không được mở file Excel đó bằng ứng dụng khác

---

## 🔄 Luồng Dữ Liệu

```
Metadata.json
    ↓
[Xác định danh sách file & template]
    ↓
File Excel Tích Hợp (01_danh_sach_chia_to)
    ↓
File Nguồn (Chấm công, Phụ cấp...)
    ↓
[Triển Khai Công Thức]
    ↓
File Excel Với Công Thức
    ↓
[User: Nhấn F9 để tính kết quả]
    ↓
File Excel Hoàn Chỉnh (Sẵn sàng Báo Cáo)
```

---

## 💡 Ví Dụ Thực Tế - Quy Trình Đầy Đủ

**Tình huống:** Công ty có 2 tổ, cần triển khai lương tháng 4/2026

### Bước 1: Tab 1 - Cấu hình cột template
- Cấu hình 12 cột (Tổ 1 → Tổ 12)
- Nhấn **"💾 ĐỒNG BỘ VÀO TEMPLATE"**

### Bước 2: Tab 2 - Tách file nhân viên
- Chọn file danh sách gốc
- Chọn Tháng 04, Năm 2026
- Nhấn **"✂️ TÁCH FILE"**
- Output:
  - `sx/Tổ Cắt - 04.xlsx` (50 NV)
  - `sx/Tổ May - 04.xlsx` (30 NV)
- Tạo metadata.json

### Bước 3: Tab 3 - Quét ảnh sản lượng
- Quét ảnh bảng sản lượng Tổ Cắt (3 ảnh, 3 mã hàng)
- Quét ảnh bảng sản lượng Tổ May (2 ảnh, 2 mã hàng)
- Output: 5 file JSON trong `02_ma_hang\2026\04\`

### Bước 4: Tab 4 - Tích hợp vào Excel
- Chế độ: "Chạy tất cả các tổ"
- Tháng 04, Năm 2026
- Nhấn **"🚀 BẮT ĐẦU TÍCH HỢP"**
- Output:
  - `Tổ Cắt - 04.xlsx` có sheet VD001, ABC-123, XYZ-456
  - `Tổ May - 04.xlsx` có sheet LMN-789, OPQ-101

### Bước 5: Tab 5 - Triển khai lương
- Chế độ: "Chạy tất cả các tổ"
- Tháng 04, Năm 2026
- Cấu hình file chấm công: `cham_cong_2026_04.xlsx`
- Nhấn **"🚀 BẮT ĐẦU TRIỂN KHAI"**
- Output:
  - Công thức lương được điền vào từng hàng nhân viên
  - Sheet "Bang TH nop" tính tổng sản lượng

### Bước 6: Xem kết quả
- Mở file Excel trong Excel
- Nhấn **F9** để tính công thức
- Xem tổng lương từng nhân viên, tổng tổ
- Sẵn sàng xuất báo cáo

---

## 🎯 Tóm Tắt Quy Trình 5 Tab

```
┌─────────────────────────────────────────────────────────┐
│ Tab 1: Cấu Hình Cột Template                             │
│ Đầu vào: Danh sách cột mới                               │
│ Đầu ra: Template cập nhật                                │
└─────────────────────────────────────────────────────────┘
                          ↓
┌─────────────────────────────────────────────────────────┐
│ Tab 2: Tách File Nhân Viên                               │
│ Đầu vào: Danh sách gốc + Tháng/Năm                      │
│ Đầu ra: File tách theo tổ + Metadata.json                │
└─────────────────────────────────────────────────────────┘
                          ↓
┌─────────────────────────────────────────────────────────┐
│ Tab 3: Quét Ảnh Sản Lượng (AI)                           │
│ Đầu vào: Ảnh bảng sản lượng                              │
│ Đầu ra: File JSON (mã hàng + công đoạn + NV)            │
└─────────────────────────────────────────────────────────┘
                          ↓
┌─────────────────────────────────────────────────────────┐
│ Tab 4: Tích Hợp Sản Lượng Vào Excel                      │
│ Đầu vào: File JSON + File Excel tách                     │
│ Đầu ra: File Excel với sheet mã hàng + Tổng hợp          │
└─────────────────────────────────────────────────────────┘
                          ↓
┌─────────────────────────────────────────────────────────┐
│ Tab 5: Triển Khai Công Thức Lương                        │
│ Đầu vào: File Excel + Metadata + Config công thức        │
│ Đầu ra: File Excel với công thức lương                   │
│ Bước cuối: Nhấn F9 để tính kết quả                       │
└─────────────────────────────────────────────────────────┘
```
