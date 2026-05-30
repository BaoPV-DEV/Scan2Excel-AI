import os
import sys
import ctypes
import subprocess
from functools import lru_cache


# =========================
# CACHE BASE PATH (GIẢM IO)
# =========================
@lru_cache(maxsize=1)
def get_base_path():
    drive = "D:/" if os.path.exists("D:/") else "C:/"
    path = os.path.join(drive, "Linh_Salary_Tool")

    os.makedirs(path, exist_ok=True)
    return path


# =========================
# SAFE CREATE DIR (FAST PATH)
# =========================
def _ensure_dir(path: str):
    if not os.path.exists(path):
        os.makedirs(path, exist_ok=True)
    return path


# =========================
# HIDE FOLDER (OPTIMIZED)
# =========================
def hide_folder(path):
    if sys.platform != "win32":
        return

    try:
        # FAST PATH (ctypes)
        if ctypes.windll.kernel32.SetFileAttributesW(path, 0x02):
            return
    except Exception:
        pass

    # FALLBACK (only if needed)
    try:
        subprocess.run(
            ["attrib", "+h", path],
            shell=True,
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
            timeout=2
        )
    except Exception:
        pass


# =========================
# CONFIG DIR (CACHE + CLEAN)
# =========================
@lru_cache(maxsize=1)
def get_config_data_dir():
    path = os.path.join(get_base_path(), "04_data")
    _ensure_dir(path)
    hide_folder(path)
    return path


# =========================
# RESOURCE PATH (PYINSTALLER SAFE)
# =========================
def resource_path(relative_path):
    base = getattr(sys, "_MEIPASS", os.path.abspath("."))
    return os.path.join(base, relative_path)


# =========================
# CONFIG FILE PATHS
# =========================
def get_salary_sources_path():
    return resource_path(os.path.join("resources", "data", "salary_sources.json"))


def get_user_run_config_path():
    return os.path.join(get_config_data_dir(), "user_run_config.json")


def get_template_columns_path():
    return os.path.join(get_config_data_dir(), "template_columns.json")


def get_metadata_path(year, month):
    month = f"{int(month):02d}"
    path = os.path.join(get_config_data_dir(), str(year), month)
    _ensure_dir(path)
    return os.path.join(path, "metadata.json")


# =========================
# DATA STRUCTURE PATHS
# =========================
@lru_cache(maxsize=128)
def get_split_excel_path(year, month):
    path = os.path.join(get_base_path(), "01_danh_sach_chia_to", str(year), str(month))
    _ensure_dir(path)
    return path


@lru_cache(maxsize=256)
def get_json_data_path(year, month, team):
    path = os.path.join(get_base_path(), "02_ma_hang", str(year), str(month), team)
    _ensure_dir(path)
    return path


# =========================
# LOG PATH (CACHE)
# =========================
@lru_cache(maxsize=1)
def get_log_path():
    path = os.path.join(get_base_path(), "03_Logs")
    _ensure_dir(path)
    return path