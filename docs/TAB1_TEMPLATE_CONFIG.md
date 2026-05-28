# Tab 1: Cập Nhật Mẫu Template - Cấu Hình Cột Tiêu Đề

## 📋 Mô Tả Chức Năng

Tab này dùng để quản lý và cấu hình các cột tiêu đề trong template Excel. Mỗi tháng, công ty có thể thay đổi danh sách tổ hoặc bộ phận nhận sản lượng, tab này cho phép bạn cấu hình các cột từ **F đến AC** (24 cột).

---

## 🎯 Quy Trình Làm Việc

### 1. **Tải Cấu Hình Hiện Tại**
- Ứng dụng tự động tải danh sách cột từ file cấu hình JSON:
  - Vị trí: `D:\Linh_Salary_Tool\04_data\template_columns.json`
- Hiển thị danh sách các cột đã được cấu hình trong bảng

### 2. **Thêm Cột Tiêu Đề Mới**
- Nhập tên tổ/bộ phận vào ô "Nhập Tên Tiêu Đề"
- Nhấn nút **"➕ Thêm Mới"** hoặc nhấn **Enter**
- Cột mới sẽ được thêm vào cuối danh sách
- Được gán vào các cột Excel từ **F (6) → AC (29)** theo thứ tự

### 3. **Sửa Tiêu Đề Cột**
- Chọn một hàng từ bảng danh sách
- Nhập tên mới vào ô "Nhập Tên Tiêu Đề"
- Nhấn nút **"📝 Sửa Tiêu Đề"**

### 4. **Xóa Tiêu Đề Cột**
- Chọn một hàng từ bảng danh sách
- Nhấn nút **"❌ Xóa Tiêu Đề"**

### 5. **Sắp Xếp Thứ Tự Cột**
- Chọn một hàng từ bảng
- Sử dụng **"🔼 Di Chuyển Lên"** hoặc **"🔽 Di Chuyển Xuống"** để thay đổi vị trí
- Cột sẽ được gán lại theo thứ tự mới

### 6. **Đặt Lại Mặc Định**
- Nhấn nút **"🔄 Đặt Lại Mặc Định"** để khôi phục danh sách 18 cột tiêu chuẩn
- Danh sách mặc định: Tổ 1 → Tổ 12, Là TP, Kiểm hoá, Cơ động, Cắt, Hoàn thiện, Tái chế

### 7. **Đồng Bộ Vào Template**
- Nhấn nút **"💾 ĐỒNG BỘ VÀO TEMPLATE"** (nút xanh)
- Ứng dụng sẽ:
  - Cập nhật 2 file template: `to_may_template.xlsx` + `kiem_hoa_template.xlsx`
  - Ẩn các cột Excel không sử dụng (từ cột sau cột cuối cùng đến AC)
  - Lưu lại cấu hình vào file JSON

---

## 📁 File Sử Dụng

| File | Vị Trí | Mục Đích |
|------|--------|---------|
| **template_columns.json** | `D:\Linh_Salary_Tool\04_data\` | Lưu trữ cấu hình cột |
| **to_may_template.xlsx** | `Template/sx/` | Template cho tổ máy |
| **kiem_hoa_template.xlsx** | `Template/gt/` | Template cho kiểm hoá/GT |

---

## ⚠️ Lưu Ý Quan Trọng

- **Không thể xóa tất cả cột**: Phải giữ ít nhất 1 cột tiêu đề
- **Giới hạn số cột**: Tối đa 24 cột (F → AC)
- **Tự động ẩn cột**: Các cột không sử dụng sẽ tự động bị ẩn trong template
- **Cập nhật ngay**: Sau khi đồng bộ, các file Excel mẫu được cập nhật lập tức
- **Dùng trước Tab 2**: Phải cấu hình cột này trước khi chạy Tab 2 (Tách File)

---

## 🔄 Luồng Dữ Liệu

```
Cấu hình Cột
    ↓
Lưu vào JSON (04_data)
    ↓
Đồng bộ vào 2 file template
    ↓
Cột được sử dụng cho Tab 2 (Tách File) & Tab 5 (Lương)
```

---

## 💡 Ví Dụ Thực Tế

Nếu công ty có 15 tổ, bạn sẽ:
1. Nhập tên từng tổ: "Tổ 1", "Tổ 2", ..., "Tổ 15"
2. Nhấn **"💾 ĐỒNG BỘ VÀO TEMPLATE"**
3. Template sẽ:
   - Dùng 15 cột: F → T (Cột 6 → Cột 20)
   - Ẩn các cột sau T (U → AC)
   - Sẵn sàng cho Tab 2 (Tách File)
