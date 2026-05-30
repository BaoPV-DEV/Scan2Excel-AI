import os
import gc
import pythoncom
import win32com.client as win32
from dotenv import load_dotenv


# =========================
# CONFIG
# =========================
START_COL = 6   # F
END_COL = 29    # AC
SHEET_NAME = "Bang TH nop"


# =========================
# UTILS
# =========================
def get_column_letter(col_idx: int) -> str:
    result = ""
    while col_idx > 0:
        col_idx, rem = divmod(col_idx - 1, 26)
        result = chr(65 + rem) + result
    return result


def safe_open_workbook(excel_app, path, password=None):
    """Mở workbook an toàn (có retry fallback không password)."""
    try:
        return excel_app.Workbooks.Open(
            os.path.abspath(path),
            UpdateLinks=0,
            ReadOnly=False,
            Password=password
        )
    except Exception:
        return excel_app.Workbooks.Open(os.path.abspath(path))


def find_sheet(wb, name: str):
    for s in wb.Sheets:
        if s.Name == name:
            return s
    return None


def set_headers(ws, titles, log_callback):
    """
    Ghi headers F -> AC + ẩn cột thừa.
    """
    ws.Range(f"F4:{get_column_letter(END_COL)}4").ClearContents()

    max_len = END_COL - START_COL + 1
    titles = titles[:max_len]

    # batch read/write COM (giảm overhead)
    for i, col in enumerate(range(START_COL, END_COL + 1)):
        col_letter = get_column_letter(col)

        if i < len(titles):
            val = titles[i].strip()
            ws.Cells(4, col).Value = val
            ws.Columns(col).Hidden = False
        else:
            ws.Columns(col).Hidden = True


# =========================
# CORE FUNCTION
# =========================
def update_templates_with_headers(titles, log_callback, progress_callback=None):
    """
    Update headers F->AC dòng 4 cho 2 template SX.
    """

    load_dotenv()
    EXCEL_PASS = os.getenv("EXCEL_SHEET_PASSWORD", "8863")

    base_dir = os.path.abspath(os.path.join(os.path.dirname(__file__), "../.."))

    templates = {
        "to_may": os.path.join(base_dir, "resources", "templates", "sx", "to_may_template.xlsx"),
        "kiem_hoa": os.path.join(base_dir, "resources", "templates", "sx", "kiem_hoa_template.xlsx"),
    }

    # validate file tồn tại
    for k, p in templates.items():
        if not os.path.exists(p):
            log_callback(f"❌ Missing template: {k} -> {p}")
            return False, f"Missing {k}"

    if len(titles) > (END_COL - START_COL + 1):
        titles = titles[: (END_COL - START_COL + 1)]
        log_callback("⚠️ Titles truncated to fit F->AC")

    pythoncom.CoInitialize()
    excel = None
    success = 0

    try:
        excel = win32.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        excel.ScreenUpdating = False

        for idx, (name, path) in enumerate(templates.items(), 1):
            log_callback(f"\n📄 Processing: {os.path.basename(path)}")

            wb = None
            try:
                wb = safe_open_workbook(excel, path, EXCEL_PASS)
                ws = find_sheet(wb, SHEET_NAME)

                if not ws:
                    log_callback(f"❌ Missing sheet: {SHEET_NAME}")
                    continue

                try:
                    ws.Unprotect(Password=EXCEL_PASS)
                except:
                    ws.Unprotect()

                set_headers(ws, titles, log_callback)

                wb.Save()
                success += 1
                log_callback(f"💾 Saved: {os.path.basename(path)}")

            except Exception as e:
                log_callback(f"❌ Error: {e}")

            finally:
                if wb:
                    try:
                        wb.Close(SaveChanges=True)
                    except:
                        pass
                    del wb

            if progress_callback:
                progress_callback(int(idx / len(templates) * 100))

        return success == len(templates), f"{success}/{len(templates)} updated"

    finally:
        if excel:
            try:
                excel.Quit()
            except:
                pass
            del excel

        pythoncom.CoUninitialize()
        gc.collect()