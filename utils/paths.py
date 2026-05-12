import os

# --- CẤU HÌNH CẤU TRÚC THƯ MỤC CHUẨN CỦA HỆ THỐNG ---
# Dữ liệu được tổ chức theo: Linh_Salary_Tool / {loại_dữ_liệu} / {năm} / {tháng} / {tổ}
def get_base_path():

    drive = "D:/" if os.path.exists("D:/") else "C:/"
    base_path = os.path.join(drive, "Linh_Salary_Tool")
    if not os.path.exists(base_path):
        os.makedirs(base_path, exist_ok=True)
    return base_path

# Trả về đường dẫn thư mục lưu trữ các file Excel đã tách theo tổ.
# Cấu trúc: Linh_Salary_Tool/01_danh_sach_chia_to/YYYY/MM
def get_split_excel_path(year, month):
    base = get_base_path()
    path = os.path.join(base, "01_danh_sach_chia_to", year, month)
    os.makedirs(path, exist_ok=True)
    return path

# Trả về đường dẫn thư mục lưu trữ các file JSON kết quả quét AI.
# Cấu trúc: Linh_Salary_Tool/02_ma_hang/YYYY/MM/Ten_To
def get_json_data_path(year, month, team):
    base = get_base_path()
    path = os.path.join(base, "02_ma_hang", year, month, team)
    os.makedirs(path, exist_ok=True)
    return path

# Trả về đường dẫn thư mục lưu trữ các logs.
# Cấu trúc: Linh_Salary_Tool/03_Logs/app.log
def get_log_path():
    base = get_base_path()
    path = os.path.join(base, "03_Logs")
    os.makedirs(path, exist_ok=True)
    return path
