# ĐẶC TẢ LOGIC HỆ THỐNG - TAB 1: PHÂN CHIA DANH SÁCH NHÂN VIÊN (DANH SÁCH CHIA TỔ)

## 1. THÔNG TIN INPUT & GIAO DIỆN (UI)
*   **Giao diện (UI):** Giữ nguyên cấu trúc UI hiện tại, không thay đổi layout hiển thị.
*   **Thành phần Input:**
    *   `File danh sách NV`: Đường dẫn chọn file Excel chứa danh sách tổng.
    *   `Checkbox chọn Tháng/Năm`: Cho phép người dùng chọn Tháng và Năm làm việc (Ví dụ: `4/2026`).
*   **Gán biến môi trường:**
    *   `INPUT_MONTH` = Tháng được chọn (Định dạng 2 chữ số, ví dụ: `04`).
    *   `INPUT_YEAR` = Năm được chọn (Định dạng 4 chữ số, ví dụ: `2026`).

---

## 2. PHÂN TÍCH FILE INPUT ("File danh sách NV")
*   **Cấu trúc dữ liệu đọc từ File danh sách NV:**
    *   Dữ liệu nhân viên bắt đầu khai báo từ **Dòng số 3**.
    *   **Cột B:** HN/TTN (Chức vụ/Vai trò - Tổ trưởng/Tổ phó/Nhân viên...)
    *   **Cột C:** Tổ bộ phận
    *   **Cột D:** Mã số NV
    *   **Cột E:** Họ và Tên

*   **Quy trình kiểm tra logic (Quy tắc loại trừ theo thứ tự ưu tiên từ trên xuống dưới):**

[Bắt đầu duyệt từ Dòng 3]
|
+---> [Bước 1: Check Thời vụ]
|     Cột D chứa "TVLA"?
|       ├── Đúng => Thống kê vào file: "Thời vụ - [INPUT_MONTH].xlsx"
|       │           Sử dụng Template: Template\sx\to_may_template.xlsx
|       └── Sai  => Chuyển xuống [Bước 2]
|
+---> [Bước 2: Check Tổ 1 -> Tổ 12]
|     Cột C thuộc nhóm ("Tổ 1", "Tổ 2", ..., "Tổ 12")?
|       ├── Đúng => Kiểm tra Cột B:
|       │             ├── B là "Tổ trưởng" => Đưa vào file: "Tổ trưởng may - [INPUT_MONTH].xlsx"
|       │             │                       Sử dụng Template: Template\sx\to_may_template.xlsx
|       │             └── B KHÔNG PHẢI tổ trưởng => Đưa vào file tương ứng: "Tổ [X].xlsx" (X từ 1-12)
|       │                                           Sử dụng Template: Template\sx\to_may_template.xlsx
|       └── Sai  => Chuyển xuống [Bước 3]
|
+---> [Bước 3: Check các tổ bộ phận đặc thù]
Duyệt giá trị Cột C để phân loại template (Đã đảm bảo KHÔNG PHẢI Thời vụ và KHÔNG PHẢI Tổ trưởng):
│
├── Cột C thuộc ("Cắt", "H.thiện", "Điều động", "Hậu giặt", "In")
│     └── Fill vào file tương ứng. Sử dụng Template: Template\sx\to_may_template.xlsx
│
├── Cột C = "Kiểm hóa"
│     └── Fill vào file tương ứng. Sử dụng Template: Template\sx\kiem_hoa_template.xlsx
│
├── Cột C = "Bảo vệ"
│     └── Fill vào file tương ứng. Sử dụng Template: Template\gt\bao_ve_template.xlsx
│
├── Cột C thuộc ("Bếp ăn", "Công vụ")
│     └── Fill vào file tương ứng. Sử dụng Template: Template\gt\bep_an_va_cong_vu_template.xlsx
│
├── Cột C = "Bốc vác"
│     └── Fill vào file tương ứng. Sử dụng Template: Template\gt\boc_vac_template.xlsx
│
├── Cột C = "C.điện"
│     └── Fill vào file tương ứng. Sử dụng Template: Template\gt\co_dien_template.xlsx
│
├── Cột C = "K.hoạch"
│     └── Fill vào file tương ứng. Sử dụng Template: Template\gt\ke_hoach_template.xlsx
│
├── Cột C = "K.thuật"
│     └── Fill vào file tương ứng. Sử dụng Template: Template\gt\ky_thuat_template.xlsx
│
├── Cột C = "VP"
│     └── Fill vào file tương ứng. Sử dụng Template: Template\gt\van_phong_template.xlsx
│
└── CÁC TRƯỜNG HỢP CÒN LẠI (Không đánh tổ, hoặc thuộc tổ: "Nghỉ không lương", "Nghỉ thai sản", "Q.trị")
└── Fill vào file: "Danh sách chưa được xử lý - [INPUT_MONTH].xlsx"
Sử dụng Template: Template\chua_xu_ly_template.xlsx


---

## 3. QUY TRÌNH XỬ LÝ CHI TIẾT TRÊN CÁC TRÊN TEMPLATE

> **Lưu ý chung cho tất cả các template có sheet "LƯƠNG Tx":**
> Khi kiểm tra và thao tác trên sheet "LƯƠNG Tx", phải chuyển đổi tên sheet về dạng viết thường (lowercase) để thực hiện kiểm tra không phân biệt hoa thường. Sau khi xử lý xong, tiến hành đổi tên sheet đó thành `"Lương " + INPUT_MONTH` (Ví dụ: `Lương 04`).

### Nhóm 3.1: `Template\sx\to_may_template.xlsx` và `Template\sx\kiem_hoa_template.xlsx`
1.  **Tạo Sheet mới:** Tạo thêm 1 sheet ở vị trí **cuối cùng** đặt tên là `"Danh sách"`.
2.  **Đổ dữ liệu vào sheet "Danh sách":** Ghi các cột dữ liệu theo thứ tự: `STT` (Cột A - Tự tăng), `HN/TTN` (Cột B), `Tổ/bộ phận` (Cột C), `Mã số NV` (Cột D), `Họ và Tên` (Cột E) lấy tương ứng từ file Input gốc.
3.  **Cập nhật thông tin sheet Lương:** 
    *   Đổi tên sheet `"LƯƠNG Tx"` thành `"Lương " + INPUT_MONTH`.
    *   Cập nhật giá trị tại ô **O3** theo định dạng: `"Tháng " + INPUT_MONTH + " năm " + INPUT_YEAR` (Ví dụ: `Tháng 04 năm 2026`).
4.  **Cập nhật thông tin sheet CĐ:**
    *   Tại ô **D3**, cập nhật định dạng: `"Tháng " + INPUT_MONTH + " năm " + INPUT_YEAR`.

### Nhóm 3.2: `Template\gt\bao_ve_template.xlsx` và `Template\gt\bep_an_va_cong_vu_template.xlsx` (Mẫu 1)
1.  **Thao tác trên sheet "Bang TH nop":**
    *   Điền thông tin nhân viên bắt đầu từ **Dòng số 9**.
    *   Cột A: STT (Tăng dần).
    *   Cột B: Ứng với cột `Họ và Tên` (Cột E file gốc).
    *   Cột C: Ứng với cột `Mã số NV` (Cột D file gốc).
    *   *Quy tắc khoảng cách:* Mỗi nhân viên cách nhau **1 dòng trống** (Ví dụ: NV1 ở dòng 9, dòng 10 bỏ trống, NV2 ở dòng 11).
2.  **Thao tác trên sheet Lương & CĐ:**
    *   Đổi tên sheet `"LƯƠNG Tx"` thành `"Lương " + INPUT_MONTH`.
    *   Cập nhật ô **N3** theo định dạng: `"Tháng " + INPUT_MONTH + " năm " + INPUT_YEAR`.
    *   Cập nhật ô **D3** tại sheet `"CĐ"` theo định dạng: `"Tháng " + INPUT_MONTH + " năm " + INPUT_YEAR`.

### Nhóm 3.3: `Template\gt\bep_an_va_cong_vu_template.xlsx` (Mẫu 2 có cập nhật tiêu đề)
1.  **Thao tác trên sheet "Bang TH nop":**
    *   Thay đổi text tại ô **A2** thành: `"BẢNG HỆ SỐ HƯỞNG LƯƠNG THÁNG " + INPUT_MONTH + "/" + INPUT_YEAR`
    *   Thay đổi text tại ô **A3** thành: `"Tổ " + [Tên tổ tương ứng]` (Ví dụ: `Tổ Bếp ăn`).
    *   Điền thông tin nhân viên bắt đầu từ **Dòng số 9** (Cột A: STT, Cột B: Họ và Tên, Cột C: Mã số NV). Cách nhau 1 dòng trống.
2.  **Thao tác trên sheet Lương & CĐ:**
    *   Đổi tên sheet `"LƯƠNG Tx"` thành `"Lương " + INPUT_MONTH`.
    *   Cập nhật ô **N3** theo định dạng: `"Tháng " + INPUT_MONTH + " năm " + INPUT_YEAR`.
    *   Cập nhật ô **D3** tại sheet `"CĐ"` theo định dạng: `"Tháng " + INPUT_MONTH + " năm " + INPUT_YEAR`.

### Nhóm 3.4: `Template\gt\boc_vac_template.xlsx`
1.  **Thao tác trên sheet "Bang TH nop":**
    *   Thay đổi ô **A2** thành: `"BẢNG HỆ SỐ HƯỞNG LƯƠNG THÁNG " + INPUT_MONTH + "/" + INPUT_YEAR`
    *   Thay đổi ô **A3** thành: `"Tổ Bốc vác"`.
    *   Điền dữ liệu từ **Dòng số 9** (Cột A: STT, Cột B: Họ Tên, Cột C: Mã NV). Cách nhau 1 dòng trống.
2.  **Thao tác trên sheet Lương & CĐ:**
    *   Đổi tên sheet thành `"Lương " + INPUT_MONTH`.
    *   Cập nhật ô **N4** theo định dạng: `"Tháng " + INPUT_MONTH + " năm " + INPUT_YEAR`.
    *   Cập nhật ô **D3** tại sheet `"CĐ"` theo định dạng: `"Tháng " + INPUT_MONTH + " năm " + INPUT_YEAR`.

### Nhóm 3.5: `Template\gt\co_dien_template.xlsx`
1.  **Thao tác trên sheet "Bang TH nop":**
    *   Thay đổi ô **A2** thành: `"BẢNG HỆ SỐ HƯỞNG LƯƠNG THÁNG " + INPUT_MONTH + "/" + INPUT_YEAR`
    *   Thay đổi ô **A3** thành: `"Tổ Cơ điện"`.
    *   Điền dữ liệu từ **Dòng số 9** (Cột A: STT, Cột B: Họ Tên, Cột C: Mã NV). Cách nhau 1 dòng trống.
2.  **Thao tác trên sheet Lương & CĐ:**
    *   Đổi tên sheet thành `"Lương " + INPUT_MONTH`.
    *   Cập nhật ô **N4** theo định dạng: `"Tháng " + INPUT_MONTH + " năm " + INPUT_YEAR`.
    *   Cập nhật ô **E3** tại sheet `"CĐ"` theo định dạng: `"Tháng " + INPUT_MONTH + " năm " + INPUT_YEAR`.

### Nhóm 3.6: `Template\gt\ke_hoach_template.xlsx` & `Template\gt\ky_thuat_template.xlsx`
1.  **Thao tác trên sheet "Bang TH nop":**
    *   Thay đổi ô **A2** thành: `"BẢNG HỆ SỐ HƯỞNG LƯƠNG THÁNG " + INPUT_MONTH + "/" + INPUT_YEAR`
    *   Thay đổi ô **A3** thành: `"Tổ " + [Tên tổ tương ứng]` (Ví dụ: `Tổ K.hoạch` hoặc `Tổ K.thuật`).
    *   Điền dữ liệu từ **Dòng số 9** (Cột A: STT, Cột B: Họ Tên, Cột C: Mã NV). Cách nhau 1 dòng trống.
2.  **Thao tác trên sheet Lương & CĐ:**
    *   Đổi tên sheet thành `"Lương " + INPUT_MONTH`.
    *   Cập nhật ô **O4** theo định dạng: `"Tháng " + INPUT_MONTH + " năm " + INPUT_YEAR`.
    *   Cập nhật ô **D3** tại sheet `"CĐ"` theo định dạng: `"Tháng " + INPUT_MONTH + " năm " + INPUT_YEAR`.

### Nhóm 3.7: `Template\gt\van_phong_template.xlsx`
1.  **Thao tác trên sheet "Bang TH nop":**
    *   Thay đổi ô **A2** thành: `"BẢNG HỆ SỐ HƯỞNG LƯƠNG THÁNG " + INPUT_MONTH + "/" + INPUT_YEAR`
    *   Thay đổi ô **A3** thành: `"Bộ phận Văn phòng"`.
    *   Điền dữ liệu từ **Dòng số 9** (Cột A: STT, Cột B: Họ Tên, Cột C: Mã NV). Cách nhau 1 dòng trống.
2.  **Thao tác trên sheet Lương & CĐ:**
    *   Đổi tên sheet thành `"Lương " + INPUT_MONTH`.
    *   Cập nhật ô **O4** theo định dạng: `"Tháng " + INPUT_MONTH + " năm " + INPUT_YEAR`.
    *   Cập nhật ô **D3** tại sheet `"CĐ"` theo định dạng: `"Tháng " + INPUT_MONTH + " năm " + INPUT_YEAR`.

---

## 4. QUY TẮC KIỂM TRA HỆ THỐNG VÀ LƯU TRỮ FILE (OUTPUT)

*   **Logic kiểm tra ổ đĩa tự động (Ẩn khỏi giao diện UI):**
    *   Hệ thống tự động quét kiểm tra sự tồn tại của ổ đĩa `D:\`.
    *   Nếu ổ `D:\` tồn tại: Gán `ROOT_PATH = "D:\Linh_Salary_Tool"`
    *   Nếu ổ `D:\` không tồn tại: Gán `ROOT_PATH = "C:\Linh_Salary_Tool"`

*   **Cấu trúc cây thư mục lưu trữ tự động:**
```text
[ROOT_PATH]
│
├── 01_danh_sach_chia_to
│   └── [INPUT_YEAR]
│       └── [INPUT_MONTH]
│           ├── Thời vụ - [INPUT_MONTH].xlsx
│           ├── Tổ trưởng may - [INPUT_MONTH].xlsx
│           ├── Tổ 1.xlsx
│           ├── Tổ 2.xlsx
│           ├── ...
│           └── Danh sách chưa được xử lý - [INPUT_MONTH].xlsx
│
├── 02_ma_hang  <--- (Thư mục chờ cho xử lý của Tab 2)
│   └── [INPUT_YEAR]
│       └── [INPUT_MONTH]
│           ├── To_1
│           └── To_2
│
└── 03_Logs     <--- (Thư mục lưu trữ nhật ký hoạt động mới)
```

5. QUY CHUẨN CÓ ĐỊNH DẠNG CODE VÀ COMMENT (DÀNH CHO ĐƠN VỊ PHÁT TRIỂN)
Tất cả các hàm, lớp (class) viết mới hoặc chỉnh sửa phải tuân thủ nghiêm ngặt quy tắc bổ sung comment bằng Tiếng Việt theo cấu trúc chuẩn hóa dưới đây:

Python
# ==============================================================================
# Thực hiện khởi tạo cấu trúc thư mục lưu trữ dựa trên việc kiểm tra ổ đĩa hệ thống
# ==============================================================================
def init_storage_path():
    # Logic xử lý tại đây
    pass

# ==============================================================================
# Thực hiện phân tích dữ liệu đầu vào từ File danh sách NV và phân chia theo case
# ==============================================================================
def process_employee_data(file_path, month, year):
    # Logic xử lý tại đây
    pass

# ==============================================================================
# Thực hiện ghi và cập nhật dữ liệu vào các File Excel dựa trên mẫu Template
# ==============================================================================
def write_data_to_template(template_path, target_path, data_list):
    # Logic xử lý tại đây
    pass