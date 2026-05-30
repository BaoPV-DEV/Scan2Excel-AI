# ==============================================================================
# SALARY SYSTEM - CONSOLIDATED PIPELINE (FIXED VERSION)
# Gồm:
#   1. Salary Config Parser (JSON template mapping)
#   2. Metadata Manager (Tab 2 output reader)
#   3. Excel COM Engine (deploy công thức)
#   4. Process pipeline (Tab 1 + Tab 5 integration)
# ==============================================================================

import os
import sys
import json
import re
from typing import Dict, List, Optional, Tuple


# ==============================================================================
# CONFIG: SALARY SOURCES
# ==============================================================================

class SalarySourceConfig:
    """
    Parse salary_sources.json:
    - mapping template → source files
    - resolve dynamic file path theo YYYY/MM
    """

    def __init__(self, config_path: str = None):
        from app.utils.paths import get_salary_sources_path, get_base_path

        self.config_path = config_path or get_salary_sources_path()
        self.config = self._load_config()
        self.base_path = self.config.get("base_path", get_base_path())
        self.templates = self.config.get("templates", {})

    def _load_config(self) -> dict:
        with open(self.config_path, "r", encoding="utf-8") as f:
            return json.load(f)

    def get_template_config(self, template_key: str) -> Optional[Dict]:
        return self.templates.get(template_key)

    def resolve_file_path(self, file_pattern: str, year: int, month: int, log_callback=None):
        """
        Resolve file pattern → real file path
        """
        yyyy = str(year)
        mm = f"{int(month):02d}"
        yy_short = yyyy[-2:]
        mm_int = int(month)

        prev_month = 12 if mm_int == 1 else mm_int - 1
        prev_year = year - 1 if mm_int == 1 else year

        resolved = file_pattern.format(
            yyyy=yyyy,
            mm=mm,
            mm_int=mm_int,
            yy_short=yy_short,
            prev_month=f"{prev_month:02d}",
            prev_year=str(prev_year),
            prev_yy_short=str(prev_year)[-2:]
        )

        full_path = os.path.join(self.base_path, yyyy, mm, resolved)

        if os.path.exists(full_path):
            return full_path

        if log_callback:
            log_callback(f"⚠️ File not found: {full_path}")
        return None


# ==============================================================================
# METADATA MANAGER (TAB 2 OUTPUT)
# ==============================================================================

class MetadataManager:
    """
    Đọc metadata.json từ Tab 2:
    - mapping file output
    - template tương ứng
    """

    def __init__(self, metadata_path: str = None):
        self.metadata_path = metadata_path
        self.metadata = {}

    def load_metadata(self, year: int, month: int) -> bool:
        from app.utils.paths import get_metadata_path

        if not self.metadata_path:
            self.metadata_path = get_metadata_path(year, month)

        if not os.path.exists(self.metadata_path):
            return False

        with open(self.metadata_path, "r", encoding="utf-8") as f:
            self.metadata = json.load(f)

        return True

    def get_split_files(self) -> Dict:
        return self.metadata.get("mapping", {})

    def get_all_groups_with_templates(self) -> List[Tuple[str, str, str]]:
        result = []
        for path, info in self.get_split_files().items():
            result.append((path, info.get("template"), info.get("folder", "root")))
        return result

    def validate_metadata(self, log_callback=None):
        errors = []

        if not self.metadata:
            return False, ["Metadata rỗng"]

        if "mapping" not in self.metadata:
            errors.append("Missing mapping")

        for path, info in self.metadata.get("mapping", {}).items():
            if "template" not in info:
                errors.append(f"Missing template: {path}")

        if log_callback:
            for e in errors:
                log_callback(e)

        return len(errors) == 0, errors


# ==============================================================================
# EXCEL COM ENGINE
# ==============================================================================

XL_CALC_MANUAL = -4135
EXCEL_PASS = os.getenv("EXCEL_SHEET_PASSWORD", "8863")


class ExcelSession:
    """
    Excel COM engine:
    - mở file
    - ghi formula
    - xử lý VLOOKUP external
    """

    def __init__(self):
        self.excel = None
        self._com = False

    # ---------------- INIT EXCEL ----------------
    def start(self, log_callback):
        if sys.platform != "win32":
            return False

        import pythoncom
        import win32com.client as win32

        pythoncom.CoInitialize()
        self._com = True

        self.excel = win32.DispatchEx("Excel.Application")
        self.excel.Visible = False
        self.excel.DisplayAlerts = False
        self.excel.Calculation = XL_CALC_MANUAL

        return True

    def stop(self):
        try:
            if self.excel:
                self.excel.Quit()
        except:
            pass
        self.excel = None

        if self._com:
            import pythoncom
            pythoncom.CoUninitialize()

    # ---------------- WRITE CELL ----------------
    def write_formula(self, ws, row, col, formula):
        try:
            cell = ws.Cells(row, col)
            cell.Formula = formula
            return True
        except:
            return False

    # ---------------- DEPLOY FILE ----------------
    def deploy_file(self, file_path, template_cfg, log_callback):
        wb = self.excel.Workbooks.Open(file_path)

        try:
            for sheet_name, cfg in template_cfg.get("sheets", {}).items():
                ws = wb.Sheets(sheet_name)

                start_row = cfg.get("start_row", 9)
                step = cfg.get("row_step", 2)

                for col, col_cfg in cfg.get("columns", {}).items():
                    col_index = self._col_to_index(col)

                    for i in range(50):
                        row = start_row + i * step
                        formula = col_cfg.get("formula", "").format(row=row)
                        self.write_formula(ws, row, col_index, formula)

            wb.Save()
            return True

        finally:
            wb.Close(False)

    def _col_to_index(self, col):
        return ord(col.upper()) - 64


# ==============================================================================
# PIPELINE TAB 1 → TAB 5
# ==============================================================================

def process_pipeline(input_path, output_dir, mm, yyyy, log_callback):
    """
    Pipeline chính:
    1. đọc file
    2. classify
    3. write excel
    """

    log_callback("🚀 START PIPELINE")

    # STEP 1 - read
    from app.logic.tab1_classifier import read_employee_data, classify_employees
    df = read_employee_data(input_path, log_callback)

    # STEP 2 - classify
    classified = classify_employees(df, mm, log_callback)

    # STEP 3 - write output files
    from app.logic.tab1_writer import write_all_groups
    success, msg = write_all_groups(classified, output_dir, mm, yyyy, log_callback)

    log_callback(msg)
    return success, msg


# ==============================================================================
# MAIN DEPLOY FROM METADATA
# ==============================================================================

def deploy_from_metadata(year, month, log_callback):
    config = SalarySourceConfig()
    metadata = MetadataManager()

    if not metadata.load_metadata(year, month):
        return False, "Missing metadata"

    session = ExcelSession()
    if not session.start(log_callback):
        return False, "Excel init failed"

    try:
        for file_path, template, _ in metadata.get_all_groups_with_templates():
            session.deploy_file(file_path, config.get_template_config(template), log_callback)

        return True, "DONE"

    finally:
        session.stop()