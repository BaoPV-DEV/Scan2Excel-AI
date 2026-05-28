# Tab 4: Tích Hợp Sản Lượng Vào Excel

## 📋 Mô Tả Chức Năng

Tab này tích hợp dữ liệu sản lượng từ các file JSON (được bóc tách ở Tab 3) vào các file Excel tách theo tổ (từ Tab 2). Tự động tạo sheet chi tiết cho từng mã hàng và cập nhật bảng tổng hợp sản lượng.

---

## 🎯 Quy Trình Làm Việc

### 1. **Chọn Chế Độ Chạy**

#### Chế Độ 1: Chạy Tất Cả Tổ
- **Radio button**: "Chạy tất cả các tổ"
- Ứng dụng sẽ quét **toàn bộ folder** `02_ma_hang\YYYY\MM\` và xử lý tất cả tổ
- Folder tháng/năm sẽ bị **vô hiệu hóa** (không chọn được)

#### Chế Độ 2: Chạy 1 Tổ
- **Radio button**: "Chạy 1 tổ"
- Nhấn **"📂 Chọn Folder"** để chỉ định folder của 1 tổ cụ thể
- Ứng dụng sẽ **chỉ xử lý tổ được chọn**
- Folder tháng/năm sẽ **vô hiệu hóa**

### 2. **Cấu Hình Thời Gian (Nếu Chế Độ Tất Cả)**
- Chọn **Tháng** (01 → 12)
- Chọn **Năm** (thường là năm hiện tại)
- Phải khớp với **Năm/Tháng ở Tab 2 & Tab 3**

### 3. **Bắt Đầu Tích Hợp**
- Nhấn nút **"🚀 BẮT ĐẦU TÍCH HỢP"** (nút xanh)
- Ứng dụng sẽ:
  - **Tìm kiếm** tất cả file JSON trong folder `02_ma_hang\YYYY\MM\{Tên Tổ}\`
  - **Mở file Excel** của tổ từ folder `01_danh_sach_chia_to\YYYY\MM\`
  - **Đối chiếu mã nhân viên** từ JSON với nhân viên trong file Excel
  - **Tạo sheet mới** cho mỗi mã hàng (VD: "VD001", "ABC-123")
  - **Điền sản lượng** vào sheet tương ứng
  - **Cập nhật bảng tổng hợp** trong sheet "Bang TH nop"

### 4. **Theo Dõi Tiến Độ**
- Thanh tiến độ hiển thị % hoàn tất
- Logs chi tiết hiển thị:
  - Số file JSON đã xử lý
  - Số sheet được tạo
  - Số nhân viên được điền dữ liệu
  - Cảnh báo hoặc lỗi (nếu có)

### 5. **Hoàn Tất**
- Thông báo **"✅ Hoàn tất!"** với số file được xử lý thành công
- File Excel được cập nhật tự động
- Sẵn sàng cho Tab 5 (Triển Khai Lương)

---

## 📁 File Sử Dụng

| Loại | File | Vị Trí |
|------|------|--------|
| **Input JSON** | Dữ liệu từ Tab 3 | `D:\Linh_Salary_Tool\02_ma_hang\YYYY\MM\{Tên Tổ}\` |
| **Input Excel** | File tách từ Tab 2 | `D:\Linh_Salary_Tool\01_danh_sach_chia_to\YYYY\MM\` |
| **Output Excel** | File Excel cập nhật | Cùng vị trí (01_danh_sach_chia_to) |

---

## 📊 Cấu Trúc File Excel Sau Tích Hợp

```
{Tên Tổ} - MM.xlsx
├── "Danh sách" (Sheet gốc)
│   ├── Tên nhân viên
│   ├── Mã nhân viên
│   └── Các tổ/bộ phận (cột D-AC)
│
├── "Bang TH nop" (Bảng tổng hợp)
│   ├── Tất cả mã hàng được tích hợp
│   └── Tổng sản lượng từ từng mã hàng
│
├── "VD001" (Sheet mã hàng)
│   ├── Công đoạn 1, 2, 3, ...
│   ├── Tên nhân viên & số lượng
│   └── Tổng sản lượng
│
├── "ABC-123" (Sheet mã hàng khác)
│   └── Tương tự như trên
│
└── ... (Thêm nhiều sheet cho mỗi mã hàng)
```

---

## 🔗 Đối Chiếu Dữ Liệu

### Quy Trình Đối Chiếu
1. **Đọc JSON** từ Tab 3 (gồm mã nhân viên, số lượng)
2. **So sánh mã nhân viên** với danh sách trong file Excel
3. **Tìm hàng tương ứng** trong sheet "Bang TH nop"
4. **Điền sản lượng** vào cột của mã hàng đó

### Xử Lý Không Tìm Thấy
- Nếu **mã nhân viên không tìm thấy** trong file Excel → **Ghi log cảnh báo**
- Dữ liệu vẫn được lưu vào sheet mã hàng (có thể kiểm tra sau)
- **Không dừng quá trình**: Tiếp tục xử lý JSON khác

---

## ⚠️ Lưu Ý Quan Trọng

### Dữ Liệu Input
- **File Excel từ Tab 2** phải tồn tại trong `01_danh_sach_chia_to\YYYY\MM\`
- **File JSON từ Tab 3** phải ở đúng folder `02_ma_hang\YYYY\MM\{Tên Tổ}\`
- **Năm/Tháng phải khớp** giữa Tab 2, 3, 4

### Mã Nhân Viên
- **Phải chính xác**: Mã nhân viên trong JSON phải khớp 100% với file Excel
- **Không phân biệt HOA/thường**: Ứng dụng so sánh không phân biệt chữ hoa/thường
- **Bỏ khoảng trắng thừa**: Ứng dụng tự động loại bỏ khoảng trắng đầu/cuối

### File Excel
- **Không được mở**: File Excel không được mở bằng ứng dụng khác trong quá trình xử lý
- **Tự động lưu**: Ứng dụng tự động lưu file sau khi cập nhật
- **Backup khuyến nghị**: Nên sao lưu file trước khi chạy

### Performance
- **Batch Processing**: Ứng dụng xử lý tất cả file một lần
- **Lặp lại được**: Có thể chạy lại mà không lo lặp lại dữ liệu

---

## 🔄 Luồng Dữ Liệu

```
File JSON (02_ma_hang)
    ↓
File Excel Tách (01_danh_sach_chia_to)
    ↓
[Đối Chiếu Mã Nhân Viên]
    ↓
[Tạo Sheet & Điền Dữ Liệu]
    ↓
[Cập Nhật Bảng Tổng Hợp]
    ↓
File Excel Cập Nhật (01_danh_sach_chia_to)
    ↓
Sử dụng ở Tab 5 (Triển Khai Lương)
```

---

## 💡 Ví Dụ Thực Tế

**Tổ Cắt có:**
- File Excel: `sx/Tổ Cắt - 04.xlsx` (50 nhân viên)
- JSON: `02_ma_hang\2026\04\Tổ Cắt\VD001.json`
  - Công đoạn Cắt: Nhân viên 001 (50 chiếc), Nhân viên 002 (30 chiếc)
- JSON: `02_ma_hang\2026\04\Tổ Cắt\ABC-123.json`
  - Công đoạn Cắt: Nhân viên 001 (40 chiếc)

**Sau tích hợp:**
- File Excel sẽ có sheet mới:
  - "VD001" → Cắt: NV001 (50), NV002 (30)
  - "ABC-123" → Cắt: NV001 (40)
- Sheet "Bang TH nop":
  - NV001: 90 chiếc (50 VD001 + 40 ABC-123)
  - NV002: 30 chiếc (VD001)
  - Tổng tổ: 120 chiếc

Lúc này file Excel sẵn sàng cho Tab 5 (Triển Khai Lương)
