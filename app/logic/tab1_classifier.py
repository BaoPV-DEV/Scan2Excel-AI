import os
import pandas as pd
import numpy as np

# ==============================================================================
# Định nghĩa bảng ánh xạ Tổ/bộ phận -> Template và tên file output
# ==============================================================================

# Đường dẫn gốc tới thư mục Template (tương đối từ thư mục gốc project)
TEMPLATE_DIR = os.path.normpath(
    os.path.join(
        os.path.dirname(os.path.abspath(__file__)), 
        "../..", 
        "resources", 
        "templates"
    )
)

# Các tổ sản xuất dùng to_may_template
TO_MAY_GROUPS = ["tổ 1","tổ 2","tổ 3","tổ 4","tổ 5","tổ 6","tổ 7","tổ 8","tổ 9","tổ 10","tổ 11","tổ 12"]
SX_EXTRA_GROUPS = ["cắt", "h.thiện", "điều động", "hậu giặt", "in"]

# Bảng tên chuẩn (canonical) - key lowercase -> tên hiển thị cố định cho filename
# Đảm bảo "Bảo vệ" và "Bảo Vệ" đều dùng cùng 1 tên file
CANONICAL_NAME = {
    "cắt": "Cắt", "h.thiện": "H.thiện", "điều động": "Điều động",
    "hậu giặt": "Hậu giặt", "in": "In",
    "kiểm hóa": "Kiểm hóa", "bảo vệ": "Bảo vệ", "bếp ăn": "Bếp ăn",
    "công vụ": "Công vụ", "bốc vác": "Bốc vác", "c.điện": "C.điện",
    "k.hoạch": "K.hoạch", "k.thuật": "K.thuật", "vp": "VP",
}

# Nhóm đặc thù GT - mỗi tên tổ ánh xạ tới 1 template riêng
GT_TEMPLATE_MAP = {
    "kiểm hóa": "sx/kiem_hoa_template.xlsx",
    "bảo vệ":   "gt/bao_ve_template.xlsx",
    "bếp ăn":   "gt/bep_an_va_cong_vu_template.xlsx",
    "công vụ":  "gt/bep_an_va_cong_vu_template.xlsx",
    "bốc vác":  "gt/boc_vac_template.xlsx",
    "c.điện":   "gt/co_dien_template.xlsx",
    "k.hoạch":  "gt/ke_hoach_template.xlsx",
    "k.thuật":  "gt/ky_thuat_template.xlsx",
    "vp":       "gt/van_phong_template.xlsx",
}

# ==============================================================================
# Cấu hình xử lý template cho từng nhóm (sheet nào, ô nào cần cập nhật)
# ==============================================================================
TEMPLATE_CONFIG = {
    "sx/to_may_template.xlsx": {
        "has_danh_sach": True, "luong_cell": "O3", "cd_cell": "D3",
    },
    "sx/kiem_hoa_template.xlsx": {
        "has_danh_sach": True, "luong_cell": "O3", "cd_cell": "D3",
    },
    "gt/bao_ve_template.xlsx": {
        "has_danh_sach": False, "luong_cell": "O3", "cd_cell": "D3",
        "update_a2": False, "update_a3": False,
    },
    "gt/bep_an_va_cong_vu_template.xlsx": {
        "has_danh_sach": False, "luong_cell": "K3", "cd_cell": "D3",
        "update_a2": True, "update_a3": True,
    },
    "gt/boc_vac_template.xlsx": {
        "has_danh_sach": False, "luong_cell": "N4", "cd_cell": "D3",
        "update_a2": True, "update_a3": True, "a3_fixed": "Tổ Bốc vác",
    },
    "gt/co_dien_template.xlsx": {
        "has_danh_sach": False, "luong_cell": "N4", "cd_cell": "E3",
        "update_a2": True, "update_a3": True, "a3_fixed": "Tổ Cơ điện",
    },
    "gt/ke_hoach_template.xlsx": {
        "has_danh_sach": False, "luong_cell": "O4", "cd_cell": "D3",
        "update_a2": True, "update_a3": True,
    },
    "gt/ky_thuat_template.xlsx": {
        "has_danh_sach": False, "luong_cell": "O4", "cd_cell": "D3",
        "update_a2": True, "update_a3": True,
    },
    "gt/van_phong_template.xlsx": {
        "has_danh_sach": False, "luong_cell": "O4", "cd_cell": "D3",
        "update_a2": True, "update_a3": True, "a3_fixed": "Bộ phận Văn phòng",
    },
}


# ==============================================================================
# Hàm đọc và phân tích file input danh sách nhân viên
# ==============================================================================
def read_employee_data(input_path, log_callback):
    """Đọc file Excel đầu vào, trả về DataFrame đã làm sạch (bắt đầu từ dòng 3)."""
    log_callback("📖 Đang đọc file danh sách nhân viên...")

    # Đọc toàn bộ dữ liệu không có header, giống code gốc
    df = pd.read_excel(input_path, header=None)

    log_callback(f"   Tổng số hàng đọc được: {len(df)}, Tổng cột: {df.shape[1]}")

    COL_B = 1  # HN/TTN
    COL_C = 2  # Tổ bộ phận
    COL_D = 3  # Mã số NV
    COL_E = 4  # Họ và Tên

    if df.shape[1] <= COL_E:
        raise ValueError(f"File input không đủ cột (cần ít nhất cột E, hiện có {df.shape[1]} cột).")

    # Bỏ 1 dòng header đầu tiên (dữ liệu bắt đầu từ dòng 2 trong Excel = index 1)
    df = df.iloc[1:].reset_index(drop=True)

    # ==========================================================================
    # BƯỚC 1: Chuẩn hóa toàn bộ dữ liệu thành chuỗi để làm sạch trước khi fill
    # ==========================================================================
    for col in [COL_B, COL_C, COL_D, COL_E]:
        df[col] = df[col].astype(str).str.strip()
        df[col] = df[col].replace(['nan', 'None', '', 'NaN', '<NA>'], pd.NA)

    # Xóa đuôi .0 ở mã NV (do pandas đọc số thành float)
    mask_d_notna = df[COL_D].notna()
    df.loc[mask_d_notna, COL_D] = df.loc[mask_d_notna, COL_D].astype(str).str.replace(r'\.0$', '', regex=True)
    df[COL_D] = df[COL_D].replace(['nan', 'None', '', '<NA>'], pd.NA)

    # Loại bỏ hàng thiếu Họ Tên (nếu cột E trống hoàn toàn -> chắc chắn không phải nhân viên)
    df = df.dropna(subset=[COL_E])

    # Loại bỏ hàng mà Họ và Tên chỉ toàn số (ví dụ số STT 382 bị lệch cột sang cột tên)
    df = df[~df[COL_E].astype(str).str.match(r'^\d+$', na=False)]

    # ==========================================================================
    # BƯỚC 2: Forward fill cột C (Tổ bộ phận) - CHỈ thực hiện trên các dòng nhân viên thực tế
    # ==========================================================================
    df[COL_C] = df[COL_C].replace(r'^\s*$', np.nan, regex=True)
    df[COL_C] = df[COL_C].ffill()

    # Loại bỏ các hàng header lặp lại trong dữ liệu
    header_mask = (
        df[COL_C].astype(str).str.lower().str.contains('tổ/bộ phận|bộ phận', na=False, regex=True) &
        df[COL_D].astype(str).str.lower().str.contains('mã', na=False)
    )
    df = df[~header_mask]

    # Loại bỏ hàng mà cột C chỉ toàn số (thường là STT bị nhầm)
    df = df[~df[COL_C].astype(str).str.match(r'^\d+$', na=False)]

    # ==========================================================================
    # QUÉT VÀ BÁO CÁO NHÂN VIÊN TRÙNG LẶP / RÁC (CHECK NHANH)
    # ==========================================================================
    seen_ids = {}
    seen_names = {}
    dup_ids = []
    dup_names = []

    for idx, row in df.iterrows():
        # Excel gốc bắt đầu từ 1, do ta bỏ 1 dòng header nên dòng Excel gốc = index + 2
        excel_row_num = idx + 2  
        ma_nv = str(row[COL_D]) if pd.notna(row[COL_D]) else ""
        ho_ten = str(row[COL_E]) if pd.notna(row[COL_E]) else ""
        ho_ten_lower = ho_ten.lower().strip()

        # Check trùng mã NV
        if ma_nv and ma_nv != "<NA>":
            if ma_nv in seen_ids:
                dup_ids.append((ma_nv, ho_ten, excel_row_num, seen_ids[ma_nv]))
            else:
                seen_ids[ma_nv] = excel_row_num

        # Check trùng họ tên
        if ho_ten_lower:
            if ho_ten_lower in seen_names:
                dup_names.append((ho_ten, ma_nv, excel_row_num, seen_names[ho_ten_lower]))
            else:
                seen_names[ho_ten_lower] = excel_row_num

    # In báo cáo ra log ngắn gọn
    unique_dups = len(dup_names)
    if unique_dups > 0:
        log_callback(f"   ⚠️ Phát hiện {unique_dups} nhân viên trùng lặp (xuất hiện nhiều lần) trong file gốc.")
        log_callback("   👉 Hệ thống sẽ TỰ ĐỘNG khử trùng (chỉ giữ lại dòng đầu tiên) để ghi vào Excel.")
    else:
        log_callback("   ✅ Không phát hiện nhân viên trùng lặp trong file gốc.")

    log_callback(f"   Tổng số nhân viên hợp lệ sau khi lọc: {len(df)}")
    return df


# ==============================================================================
# Hàm phân loại nhân viên theo quy trình kiểm tra logic tab1.md (Mục 2)
# ==============================================================================
def classify_employees(df, mm, log_callback):
    """
    Phân loại từng nhân viên theo thứ tự ưu tiên và khử trùng toàn cục:
    Bước 1: Thời vụ (TVLA) -> Bước 2: Tổ 1-12 (check Tổ trưởng) -> Bước 3: Đặc thù
    """
    COL_B = 1  # HN/TTN
    COL_C = 2  # Tổ bộ phận
    COL_D = 3  # Mã số NV
    COL_E = 4  # Họ và Tên

    result = {}  # key = output_filename, value = {template, employees, to_name, config_key}
    processed_keys = set()  # Set lưu trữ khóa nhận diện nhân viên đã xử lý để khử trùng

    def add_to_group(filename, template_key, row_data, display_name):
        if filename not in result:
            result[filename] = {
                "template_key": template_key,
                "template_path": os.path.join(TEMPLATE_DIR, template_key),
                "employees": [],
                "to_name": display_name,
            }
        result[filename]["employees"].append(row_data)

    for _, row in df.iterrows():
        col_b = str(row[COL_B]) if pd.notna(row[COL_B]) else ""
        col_c = str(row[COL_C]) if pd.notna(row[COL_C]) else ""
        col_d = str(row[COL_D]) if pd.notna(row[COL_D]) else ""
        col_e = str(row[COL_E]) if pd.notna(row[COL_E]) else ""
        col_c_lower = col_c.lower().strip()
        col_b_lower = col_b.lower().strip()

        # ======================================================================
        # 1. KHỬ TRÙNG TOÀN CỤC (Deduplication)
        # ======================================================================
        # Tạo khóa nhận diện duy nhất
        if col_d and col_d != "<NA>":
            emp_key = f"id_{col_d.lower().strip()}"
        else:
            emp_key = f"name_{col_e.lower().strip()}"

        if emp_key in processed_keys:
            # Bỏ qua nếu nhân viên này đã được xếp vào file nào đó trước đó
            continue
        processed_keys.add(emp_key)

        emp = {"hn_ttn": col_b, "to_bp": col_c, "ma_nv": col_d, "ho_ten": col_e}

        # Check điều kiện đặc biệt của cột B
        is_totruong = col_b_lower == "tổ trưởng"
        is_qtri = "q.trị" in col_b_lower or "q.tri" in col_b_lower

        # ======================================================================
        # Bước 1: Check Thời vụ - Cột D chứa "TVLA"
        # ======================================================================
        if "tvla" in col_d.lower():
            fname = f"Thời vụ - {mm}.xlsx"
            add_to_group(fname, "sx/to_may_template.xlsx", emp, "Thời vụ")
            continue

        # ======================================================================
        # Bước 2: Check Tổ 1 -> Tổ 12 hoặc tổ Điều động (nếu là Tổ trưởng)
        # ======================================================================
        if col_c_lower in TO_MAY_GROUPS or (col_c_lower == "điều động" and is_totruong):
            if is_totruong:
                fname = f"Tổ trưởng may - {mm}.xlsx"
                add_to_group(fname, "sx/to_may_template.xlsx", emp, "Tổ trưởng may")
            elif is_qtri:
                # Nếu là Q.trị thì không cho vào file tổ bộ phận, rơi xuống cuối (chưa xử lý)
                pass
            else:
                # Lấy số tổ
                to_num = col_c_lower.replace("tổ ", "").strip()
                fname = f"Tổ {to_num} - {mm}.xlsx"
                add_to_group(fname, "sx/to_may_template.xlsx", emp, col_c)
                continue
            if is_totruong:
                continue

        # ======================================================================
        # Bước 3: Check các tổ bộ phận đặc thù
        # Điều kiện: Chỉ fill khi cột B KHÔNG PHẢI "Tổ trưởng" hoặc "Q.trị"
        # ======================================================================
        if not is_totruong and not is_qtri:
            if col_c_lower in SX_EXTRA_GROUPS:
                canon = CANONICAL_NAME[col_c_lower]
                fname = f"{canon} - {mm}.xlsx"
                add_to_group(fname, "sx/to_may_template.xlsx", emp, canon)
                continue

            if col_c_lower in GT_TEMPLATE_MAP:
                template_key = GT_TEMPLATE_MAP[col_c_lower]
                canon = CANONICAL_NAME[col_c_lower]
                fname = f"{canon} - {mm}.xlsx"
                add_to_group(fname, template_key, emp, canon)
                continue

        # ======================================================================
        # Trường hợp còn lại -> Chưa xử lý (gồm Q.trị, Tổ trưởng bộ phận ngoài may, v.v.)
        # ======================================================================
        fname = f"Danh sách chưa được xử lý - {mm}.xlsx"
        add_to_group(fname, "chua_xu_ly_template.xlsx", emp, "Chưa xử lý")

    # Log tóm tắt kết quả phân loại (không in chi tiết từng người để tránh làm chậm UI)
    log_callback(f"\n📊 Kết quả phân loại:")
    for fname, info in result.items():
        log_callback(f"   - {fname}: {len(info['employees'])} nhân viên")

    return result
