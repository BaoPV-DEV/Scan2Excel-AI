# Tab 3: Quét Ảnh Sản Lượng Bằng AI (Gemini)

## 📋 Mô Tả Chức Năng

Tab này sử dụng **Google Gemini 2.0 Flash AI** để tự động bóc tách dữ liệu từ ảnh chụp bảng sản lượng. Nhận diện chính xác: mã hàng, công đoạn, định mức, người thực hiện, số lượng. Cho phép xem và chỉnh sửa trực tiếp trước khi lưu JSON.

---

## 🎯 Quy Trình Làm Việc

### 1. **Chọn Tổ & Cấu Hình**

#### Chọn Tổ Quét
- Mở danh sách Tổ (các tổ từ Tab 2 vừa tách)
- Nếu không chọn, click **"📂 Chọn Folder Tổ"** để chỉ định thủ công
- Hoặc để ứng dụng **tự động quét** folder `D:\Linh_Salary_Tool\02_ma_hang\YYYY\MM\`

#### Cấu Hình Năm/Tháng
- Chọn **Tháng** và **Năm** (phải khớp với Tab 2)
- Output sẽ lưu tại: `D:\Linh_Salary_Tool\02_ma_hang\YYYY\MM\{Tên Tổ}\`

#### API Key (Nếu cần)
- Mặc định: Đọc từ file `.env` (biến `GEMINI_API_KEY`)
- Có thể nhập API key trực tiếp nếu file `.env` chưa cấu hình

### 2. **Chọn Ảnh/Folder Ảnh**
- Nhấn **"🖼️ Chọn Ảnh"** để chọn 1 ảnh duy nhất
- HOẶC nhấn **"📂 Chọn Folder Ảnh"** để quét **tất cả ảnh** trong folder
- Định dạng hỗ trợ: `.jpg`, `.png`, `.jpeg`, `.webp`

### 3. **Bắt Đầu Quét**
- Nhấn nút **"🤖 BẮT ĐẦU QUÉT"** (nút xanh)
- Ứng dụng sẽ:
  - **Gửi ảnh** tới Google Gemini API
  - **Bóc tách dữ liệu** theo format JSON chuẩn
  - **Tạo file JSON** tạm thời để xem/sửa

### 4. **Xem & Sửa Dữ Liệu Trực Tiếp**
- Sau khi AI bóc tách, hiển thị hộp thoại **"SỬA DỮ LIỆU SẢN XUẤT (TRỰC TIẾP)"**
- Có thể:
  - **Sửa tên, mô tả công đoạn**
  - **Sửa định mức, số lượng**
  - **Thêm công đoạn mới** (chèn vào vị trí bất kỳ)
  - **Xóa công đoạn**
  - **Chỉnh sửa danh sách nhân viên** của từng công đoạn

### 5. **Lưu JSON**
- Nhấn **"💾 LƯU"** để lưu dữ liệu đã chỉnh sửa
- File JSON được lưu tại: `D:\Linh_Salary_Tool\02_ma_hang\YYYY\MM\{Tên Tổ}\{Mã Hàng}.json`
- Nếu nhấn **"❌ HỦY"**, dữ liệu sẽ không được lưu

### 6. **Theo Dõi Tiến Độ**
- Thanh tiến độ hiển thị % hoàn tất
- Logs chi tiết hiển thị:
  - Số ảnh đã xử lý
  - Mã hàng phát hiện
  - Số công đoạn bóc tách
  - Cảnh báo (nếu có)

---

## 📁 File Sử Dụng

| Loại | File | Vị Trí |
|------|------|--------|
| **Input** | Ảnh bảng sản lượng | Bất kỳ (do user chọn) |
| **Config** | .env (GEMINI_API_KEY) | Gốc dự án |
| **Output** | JSON dữ liệu bóc tách | `D:\Linh_Salary_Tool\02_ma_hang\YYYY\MM\{Tên Tổ}\` |

---

## 📋 Cấu Trúc JSON Output

```json
{
  "thong_tin_chung": {
    "ma_hang": "VD001",
    "to_san_xuat": "Tổ Cắt",
    "thoi_gian": "15/04/2026",
    "khach_hang": "ABCD Co.",
    "loai_san_pham": "Áo Sơ Mi",
    "don_vi": "Chiếc",
    "thuong_ma_hang_moi": 0,
    "tong_san_luong_muc_tieu": 100
  },
  "danh_sach_cong_doan": [
    {
      "stt": 1,
      "mo_ta": "Cắt vải",
      "dinh_muc_t": 2.5,
      "thuc_hien": [
        {
          "nhan_vien": "Nguyễn Văn A",
          "ma_nhan_vien": "001",
          "so_luong": 50
        }
      ],
      "tong_thuc_hien": 50
    }
  ]
}
```

---

## 🤖 Khả Năng AI (Gemini)

### Nhận Diện
- ✅ Mã hàng (VD: VD001, ABC-123)
- ✅ Tên công đoạn (VD: Cắt, May, Hoàn thiện)
- ✅ Định mức thời gian (VD: 2.5 phút/chiếc)
- ✅ Tên nhân viên (từ chữ viết tay)
- ✅ Số lượng sản phẩm
- ✅ Ghi chú phức tạp (VD: `5144 - 204 = 4940`)

### Hỗ Trợ
- Ảnh chất lượng thấp, mờ, xoay
- Ảnh chứa nhiều bảng
- Bảng có ghi chú viết tay
- Ảnh in đen trắng

---

## ⚠️ Lưu Ý Quan Trọng

### API & Network
- **Cần Internet**: Bắt buộc phải kết nối Internet
- **API Key bắt buộc**: Phải cấu hình Google Gemini API Key
- **Rate limit**: Nếu quét quá nhiều ảnh cùng lúc, có thể gặp giới hạn API

### Chất Lượng Ảnh
- **Ảnh rõ ràng**: AI sẽ bóc tách chính xác hơn
- **Ánh sáng đủ**: Tránh ảnh bị đen hoặc quá sáng
- **Không bị xoay**: Ảnh nên theo chiều ngang hoặc dọc

### Dữ Liệu AI
- **Luôn kiểm tra**: Dữ liệu từ AI có thể sai, phải xem/sửa trước lưu
- **Đặc biệt chú ý**: Số lượng, mã hàng, tên nhân viên
- **Sửa ngay**: Nếu phát hiện sai, sửa trên giao diện trước lưu

---

## 🔄 Luồng Dữ Liệu

```
Ảnh Bảng Sản Lượng
    ↓
[Gửi tới Google Gemini]
    ↓
AI Bóc Tách Dữ Liệu
    ↓
[User Xem & Sửa]
    ↓
Lưu JSON (02_ma_hang)
    ↓
Sử dụng ở Tab 4 (Tích Hợp)
```

---

## 💡 Ví Dụ Thực Tế

**Ảnh bảng sản lượng có:**
- Mã hàng: VD001
- Công đoạn: Cắt (2.5 min), May (5 min), Hoàn thiện (1.5 min)
- Nhân viên: Nguyễn Văn A (50 chiếc), Trần Thị B (30 chiếc)

**Output JSON sẽ có:**
- 3 công đoạn (Cắt, May, Hoàn thiện)
- Danh sách nhân viên và số lượng họ làm
- Định mức thời gian cho mỗi công đoạn

**Bạn có thể sửa:**
- Thêm công đoạn mới giữa May và Hoàn thiện
- Sửa số lượng của nhân viên
- Sửa tên công đoạn nếu AI nhận diện sai
