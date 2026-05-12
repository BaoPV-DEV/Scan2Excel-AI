# Scan2Excel AI - Công cụ hỗ trợ sản xuất

**Scan2Excel AI** là ứng dụng desktop mạnh mẽ giúp tự động hóa quy trình quản lý sản lượng ngành may. Ứng dụng kết hợp sức mạnh của **Google Gemini AI** để bóc tách dữ liệu từ ảnh chụp bảng sản lượng và tích hợp trực tiếp vào hệ thống file Excel báo cáo.

---

## 🚀 Tính năng chính

### 1. Tách file nhân viên theo tổ (Tab 1)
- Đọc danh sách nhân viên tổng của nhà máy.
- Tự động lọc và tách thành từng file Excel riêng biệt cho mỗi tổ/bộ phận.
- Sử dụng file mẫu (Template) chuẩn để đảm bảo định dạng đồng nhất.

### 2. Quét ảnh bảng sản lượng bằng AI (Tab 2)
- Sử dụng mô hình **Gemini 2.0 Flash** để bóc tách dữ liệu.
- Nhận diện chính xác: Mã hàng, tên công đoạn, định mức thời gian, người thực hiện và số lượng sản phẩm.
- Hỗ trợ xử lý các ghi chú viết tay phức tạp (ví dụ: `5144 - 204 = 4940`).
- Cho phép xem và chỉnh sửa dữ liệu trực tiếp trước khi lưu thành file JSON.

### 3. Tích hợp dữ liệu vào Excel (Tab 3)
- Batch Process: Tự động quét hàng loạt file JSON trong thư mục.
- Đối chiếu mã nhân viên và điền sản lượng vào đúng file Excel của tổ tương ứng.
- Tự động tạo sheet chi tiết cho từng mã hàng.
- Cập nhật bảng tổng hợp lương và sản lượng nộp hàng tháng.

---

## 🛠️ Công nghệ sử dụng

- **Ngôn ngữ**: Python 3.10+
- **Giao diện**: PySide6 (Qt for Python)
- **Xử lý dữ liệu**: Pandas, OpenPyXL
- **Tương tác Excel**: win32com (đảm bảo giữ nguyên định dạng file phức tạp)
- **AI Engine**: Google GenAI (Gemini API)

---

## 📂 Cấu trúc thư mục dự án

```text
Scan2Excel-AI/
├── main.py                # Điểm khởi đầu của ứng dụng
├── logic/                 # Các module xử lý nghiệp vụ chính
│   ├── excel_processor.py   # Tách file nhân viên (Tab 1)
│   └── excel_integration.py # Tích hợp dữ liệu vào Excel (Tab 3)
├── ui/                    # Các module giao diện người dùng
│   ├── gui.py              # Cửa sổ chính và quản lý Tab
│   ├── ui_part1_split.py   # Giao diện Tab 1
│   ├── ui_part2_scan.py    # Giao diện Tab 2 (Xử lý AI)
│   └── ui_part3_link.py    # Giao diện Tab 3 (Tích hợp)
├── utils/                 # Các tiện ích hệ thống
│   ├── paths.py           # Quản lý đường dẫn dữ liệu
│   ├── logger.py          # Hệ thống ghi nhật ký
│   └── config.py          # Cấu hình ứng dụng
├── Template/              # Chứa file Excel mẫu (.xlsx)
└── scratch/               # Các script kiểm tra và nháp
```

---

## ⚙️ Cài đặt và Cấu hình

1. **Cài đặt thư viện**:
   ```bash
   pip install PySide6 pandas openpyxl python-dotenv google-genai pywin32
   ```

2. **Cấu hình biến môi trường**:
   Tạo file `.env` tại thư mục gốc với nội dung:
   ```text
   GEMINI_API_KEY=your_google_gemini_api_key_here
   EXCEL_SHEET_PASSWORD=8863
   ```

3. **Chạy ứng dụng**:
   ```bash
   python main.py
   ```

---

## 📝 Lưu ý quan trọng
- Cần có kết nối Internet để sử dụng tính năng Quét ảnh AI (Tab 2).
- Các file Excel mẫu trong thư mục `Template/` phải được giữ nguyên tên và cấu trúc sheet để ứng dụng hoạt động chính xác.
- Khi chạy ứng dụng, hãy đảm bảo không có file Excel mục tiêu nào đang được mở để tránh lỗi tranh chấp quyền truy cập.