import os
import pandas as pd
import numpy as np

# ==============================================================================
# TEMPLATE CONFIG
# ==============================================================================

BASE_DIR = os.path.abspath(os.path.join(os.path.dirname(__file__), "../.."))

TEMPLATE_DIR = os.path.join(BASE_DIR, "resources", "templates")

TO_MAY_GROUPS = [f"tổ {i}" for i in range(1, 13)]
SX_EXTRA_GROUPS = ["cắt", "h.thiện", "điều động", "hậu giặt", "in"]

CANONICAL_NAME = {
    "cắt": "Cắt",
    "h.thiện": "H.thiện",
    "điều động": "Điều động",
    "hậu giặt": "Hậu giặt",
    "in": "In",
    "kiểm hóa": "Kiểm hóa",
    "bảo vệ": "Bảo vệ",
    "bếp ăn": "Bếp ăn",
    "công vụ": "Công vụ",
    "bốc vác": "Bốc vác",
    "c.điện": "C.điện",
    "k.hoạch": "K.hoạch",
    "k.thuật": "K.thuật",
    "vp": "VP",
}

GT_TEMPLATE_MAP = {
    "kiểm hóa": "sx/kiem_hoa_template.xlsx",
    "bảo vệ": "gt/bao_ve_template.xlsx",
    "bếp ăn": "gt/bep_an_va_cong_vu_template.xlsx",
    "công vụ": "gt/bep_an_va_cong_vu_template.xlsx",
    "bốc vác": "gt/boc_vac_template.xlsx",
    "c.điện": "gt/co_dien_template.xlsx",
    "k.hoạch": "gt/ke_hoach_template.xlsx",
    "k.thuật": "gt/ky_thuat_template.xlsx",
    "vp": "gt/van_phong_template.xlsx",
}


# ==============================================================================
# SAFE STRING HELPER (FIX CRASH NA / pd.NA)
# ==============================================================================

def safe_str(v):
    if pd.isna(v):
        return ""
    return str(v)


# ==============================================================================
# READ DATA
# ==============================================================================

def read_employee_data(input_path, log_callback):
    log_callback("📖 Đọc file nhân viên...")

    df = pd.read_excel(input_path, header=None)

    COL_B, COL_C, COL_D, COL_E = 1, 2, 3, 4

    if df.shape[1] <= COL_E:
        raise ValueError("File thiếu cột dữ liệu")

    df = df.iloc[1:].reset_index(drop=True)

    for col in [COL_B, COL_C, COL_D, COL_E]:
        df[col] = df[col].astype(str).str.strip()
        df[col] = df[col].replace(['nan', 'None', '', '<NA>'], pd.NA)

    mask = df[COL_D].notna()
    df.loc[mask, COL_D] = df.loc[mask, COL_D].astype(str).str.replace(r"\.0$", "", regex=True)

    df = df.dropna(subset=[COL_E])
    df = df[~df[COL_E].astype(str).str.match(r"^\d+$", na=False)]

    df[COL_C] = df[COL_C].replace(r"^\s*$", np.nan, regex=True)
    df[COL_C] = df[COL_C].ffill()

    df = df[~df[COL_C].astype(str).str.contains("tổ/bộ phận|bộ phận", na=False)]

    log_callback(f"   Tổng NV hợp lệ: {len(df)}")
    return df


# ==============================================================================
# CLASSIFY (FIXED)
# ==============================================================================

def classify_employees(df, mm, log_callback):
    COL_B, COL_C, COL_D, COL_E = 1, 2, 3, 4

    result = {}
    processed = set()

    def add_group(fname, template, emp, name):
        if fname not in result:
            result[fname] = {
                "template_key": template,
                "template_path": os.path.join(TEMPLATE_DIR, template),
                "employees": [],
                "to_name": name,
            }
        result[fname]["employees"].append(emp)

    for _, row in df.iterrows():

        # =========================
        # FIX CRASH pd.NA HERE
        # =========================
        col_b = safe_str(row[COL_B]).strip().lower()
        col_c = safe_str(row[COL_C]).strip().lower()
        col_d = safe_str(row[COL_D]).strip()
        col_e = safe_str(row[COL_E]).strip()

        if not col_e:
            continue

        emp_key = col_d.lower() if col_d else col_e.lower()
        if emp_key in processed:
            continue
        processed.add(emp_key)

        emp = {
            "hn_ttn": col_b,
            "to_bp": col_c,
            "ma_nv": col_d,
            "ho_ten": col_e
        }

        is_truong = "tổ trưởng" in col_b
        is_qtri = "q.trị" in col_b or "q.tri" in col_b

        # ======================================================
        # 1. THỜI VỤ
        # ======================================================
        if col_d and "tvla" in col_d.lower():
            add_group(f"Thời vụ - {mm}.xlsx", "sx/to_may_template.xlsx", emp, "Thời vụ")
            continue

        # ======================================================
        # 2. TỔ MAY
        # ======================================================
        if col_c in TO_MAY_GROUPS:
            if is_truong:
                add_group(f"Tổ trưởng may - {mm}.xlsx", "sx/to_may_template.xlsx", emp, "Tổ trưởng")
            else:
                to_num = col_c.replace("tổ", "").strip()
                add_group(f"Tổ {to_num} - {mm}.xlsx", "sx/to_may_template.xlsx", emp, col_c)
            continue

        # ======================================================
        # 3. SX EXTRA + GT TEMPLATE
        # ======================================================
        if not is_truong and not is_qtri:

            if col_c in SX_EXTRA_GROUPS:
                name = CANONICAL_NAME.get(col_c, col_c)
                add_group(f"{name} - {mm}.xlsx", "sx/to_may_template.xlsx", emp, name)
                continue

            if col_c in GT_TEMPLATE_MAP:
                tpl = GT_TEMPLATE_MAP[col_c]
                name = CANONICAL_NAME.get(col_c, col_c)
                add_group(f"{name} - {mm}.xlsx", tpl, emp, name)
                continue

        # ======================================================
        # DEFAULT
        # ======================================================
        add_group(
            f"Danh sách chưa xử lý - {mm}.xlsx",
            "chua_xu_ly_template.xlsx",
            emp,
            "Chưa xử lý"
        )

    log_callback("\n📊 Kết quả phân loại:")
    for k, v in result.items():
        log_callback(f"   - {k}: {len(v['employees'])}")

    return result