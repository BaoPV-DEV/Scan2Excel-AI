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
    return os.path.join("Template","data", "salary_sources.json")

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

# Thực hiện di chuyển cấu hình cũ sang thư mục ẩn và xóa file cũ
def migrate_configs_to_hidden_dir():
    config_dir = get_config_data_dir()
    app_root = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
    
    migrations = {
        "salary_sources.json": [
            os.path.join(app_root, "Template", "salary_sources.json"),
            os.path.join(app_root, "salary_sources.json")
        ],
        "user_run_config.json": [
            os.path.join(app_root, "Template", "user_run_config.json"),
            os.path.join(app_root, "user_run_config.json")
        ],
        "template_columns.json": [
            os.path.join(config_dir, "sx", "template_columns.json"),  # Cũ dưới 04_data/sx/
            os.path.join(app_root, "Template", "template_columns.json"),
            os.path.join(app_root, "template_columns.json")
        ]
    }
    
    for target_name, src_paths in migrations.items():
        target_path = os.path.join(config_dir, target_name)
        
        found_src = None
        for src in src_paths:
            if os.path.exists(src) and os.path.abspath(src) != os.path.abspath(target_path):
                found_src = src
                break
                
        if found_src:
            # Nếu target chưa tồn tại, copy từ nguồn sang target
            if not os.path.exists(target_path):
                try:
                    os.makedirs(os.path.dirname(target_path), exist_ok=True)
                    import shutil
                    shutil.copy2(found_src, target_path)
                    print(f"[Paths] Copied config: {found_src} -> {target_path}")
                except Exception as e:
                    print(f"[Paths] Error copying {found_src} to {target_path}: {e}")
            
            # Nếu target đã tồn tại (hoặc vừa copy xong), xóa file gốc ở source
            if os.path.exists(target_path):
                try:
                    os.remove(found_src)
                    print(f"[Paths] Deleted old source config: {found_src}")
                except Exception as e:
                    print(f"[Paths] Error deleting old config {found_src}: {e}")
                    
    # Dọn dẹp thư mục sx cũ nếu trống
    try:
        sx_dir = os.path.join(config_dir, "sx")
        if os.path.exists(sx_dir) and not os.listdir(sx_dir):
            os.rmdir(sx_dir)
            print(f"[Paths] Cleaned up empty dir: {sx_dir}")
    except Exception:
        pass

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
