import os
import sys

# --- CẤU HÌNH CẤU TRÚC THƯ MỤC CHUẨN CỦA HỆ THỐNG ---
# Dữ liệu được tổ chức theo: Linh_Salary_Tool / {loại_dữ_liệu} / {năm} / {tháng} / {tổ}
def get_base_path():
    drive = "D:/" if os.path.exists("D:/") else "C:/"
    base_path = os.path.join(drive, "Linh_Salary_Tool")
    if not os.path.exists(base_path):
        os.makedirs(base_path, exist_ok=True)
    return base_path

# Ẩn thư mục trên Windows
def hide_folder(path):
    if sys.platform == 'win32':
        try:
            import ctypes
            # FILE_ATTRIBUTE_HIDDEN = 0x02
            ret = ctypes.windll.kernel32.SetFileAttributesW(path, 0x02)
            if not ret:
                import subprocess
                subprocess.run(['attrib', '+h', path], shell=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        except Exception:
            try:
                import subprocess
                subprocess.run(['attrib', '+h', path], shell=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
            except Exception:
                pass

# Trả về đường dẫn thư mục cấu hình và đánh dấu ẩn
def get_config_data_dir():
    base = get_base_path()
    path = os.path.join(base, "04_data")
    if not os.path.exists(path):
        os.makedirs(path, exist_ok=True)
    hide_folder(path)
    return path

# Trả về đường dẫn file cấu hình salary_sources.json
def get_salary_sources_path():
    return resource_path(os.path.join("resources", "data", "salary_sources.json"))

# Trả về đường dẫn file cấu hình user_run_config.json
def get_user_run_config_path():
    return os.path.join(get_config_data_dir(), "user_run_config.json")

# Trả về đường dẫn file cấu hình template_columns.json
def get_template_columns_path():
    return os.path.join(get_config_data_dir(), "template_columns.json")

# Trả về đường dẫn file cấu hình metadata.json
def get_metadata_path(year, month):
    # Đảm bảo month dạng 2 chữ số (VD: "04")
    month_str = f"{int(month):02d}"
    year_str = str(year)
    path = os.path.join(get_config_data_dir(), year_str, month_str)
    os.makedirs(path, exist_ok=True)
    return os.path.join(path, "metadata.json")


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

def resource_path(relative_path):
    if hasattr(sys, "_MEIPASS"):
        return os.path.join(sys._MEIPASS, relative_path)
    return os.path.join(os.path.abspath("."), relative_path)