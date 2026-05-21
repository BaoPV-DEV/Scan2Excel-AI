import os
import shutil
import re
import pythoncom
import win32com.client as win32
from logic.tab1_classifier import TEMPLATE_CONFIG

# ==============================================================================
# CẤU HÌNH CÔNG THỨC CHO CÁC TEMPLATE ĐẶC BIỆT
# Người dùng có thể dễ dàng thêm mới hoặc chỉnh sửa cấu hình cho các template khác tại đây.
# Các placeholder hỗ trợ:
#   {row}: Dòng hiện tại (VD: 9)
#   {row_minus_1}: Dòng hiện tại trừ 1 (VD: 8)
#   {row_plus_1}: Dòng hiện tại cộng 1 (VD: 10)
#   {yyyy}: Năm hiện tại (VD: 2026)
#   {yy_short}: Hai chữ số cuối của năm hiện tại (VD: 26)
#   {mm}: Tháng hiện tại dạng 2 chữ số (VD: 04)
#   {mm_int}: Tháng hiện tại dạng số nguyên (VD: 4)
#   {prev_year}: Năm trước (VD: 2026)
#   {prev_yy_short}: Hai chữ số cuối của năm trước (VD: 26)
#   {prev_month}: Tháng trước dạng 2 chữ số (VD: 03)
#   {prev_month_int}: Tháng trước dạng số nguyên (VD: 3)
# ==============================================================================
TEMPLATE_FORMULA_CONFIG = {
    "gt/bao_ve_template.xlsx": {
        "Bang TH nop": {
            "start_row": 9,
            "row_step": 2,
            "formulas": {
                "D": "=+IFERROR(VLOOKUP(C{row},'D:\\HỒ SƠ+LƯƠNG SD\\LƯƠNG\\LƯƠNG + CÔNG {yyyy}\\Tháng {mm}-{yyyy}\\Danh sách+Chấm công\\[Danh sách CBCNV bản dùng làm lương T{mm_int}.xlsx]Danh sách'!$D$1:$AS$1000,42,0),0)"
            }
        },
        "Lương {mm}": {
            "start_row": 10,
            "row_step": 2,
            "formulas": {
                "Z": "=+IFERROR(VLOOKUP(C{row_minus_1},'D:\\HỒ SƠ+LƯƠNG SD\\LƯƠNG\\LƯƠNG + CÔNG {prev_year}\\Tháng {prev_month}-{prev_year}\\[CK tháng {prev_month}.{prev_year}.xlsx]Bản gốc T{prev_month}'!$C$1:$J$1000,8,0),0)",
                "AA": "=+IFERROR(VLOOKUP(C{row_minus_1},'D:\\HỒ SƠ+LƯƠNG SD\\LƯƠNG\\LƯƠNG + CÔNG {prev_year}\\Tháng {prev_month}-{prev_year}\\[CK tháng {prev_month}.{prev_year}.xlsx]Bản gốc T{prev_month}'!$C$1:$F$1000,4,0),0)"
            }
        }
    },
    "gt/bep_an_va_cong_vu_template.xlsx": {
        "Bang TH nop": {
            "start_row": 9,
            "row_step": 2,
            "formulas": {
                "D": "=+IFERROR(VLOOKUP(C{row},'D:\\HỒ SƠ+LƯƠNG SD\\LƯƠNG\\LƯƠNG + CÔNG {yyyy}\\Tháng {mm}-{yyyy}\\Danh sách+Chấm công\\[Danh sách CBCNV bản dùng làm lương T{mm_int}.xlsx]Danh sách'!$D$1:$AS$1000,42,0),0)",
                "E": "=+IFERROR(VLOOKUP(C{row},'D:\\HỒ SƠ+LƯƠNG SD\\LƯƠNG\\LƯƠNG + CÔNG {yyyy}\\Tháng {mm}-{yyyy}\\Danh sách+Chấm công\\[Danh sách CBCNV làm lương tháng {mm}.xlsx]Danh Sách dùng'!$E$5:$AG$11000,29,0),0)",
                "G": "=+IFERROR(VLOOKUP(C{row},'D:\\HỒ SƠ+LƯƠNG SD\\LƯƠNG\\LƯƠNG + CÔNG {yyyy}\\Tháng {mm}-{yyyy}\\Danh sách+Chấm công\\[Chấm công {mm}.{yy_short}.xlsx]Công'!$C$5:$AO$6000,39,0),0)",
                "H": "=+IFERROR(VLOOKUP(C{row},'D:\\HỒ SƠ+LƯƠNG SD\\LƯƠNG\\LƯƠNG + CÔNG {yyyy}\\Tháng {mm}-{yyyy}\\Danh sách+Chấm công\\[Chấm công {mm}.{yy_short}.xlsx]Công'!$C$5:$AU$6000,44,0),0)",
                "I": "=+IFERROR(VLOOKUP(C{row},'D:\\HỒ SƠ+LƯƠNG SD\\LƯƠNG\\LƯƠNG + CÔNG {yyyy}\\Tháng {mm}-{yyyy}\\Danh sách+Chấm công\\[Chấm công {mm}.{yy_short}.xlsx]Công'!$C$5:$AR$6000,42,0),0)",
                "K": "=+IFERROR(VLOOKUP(C{row},'D:\\HỒ SƠ+LƯƠNG SD\\LƯƠNG\\LƯƠNG + CÔNG {yyyy}\\Tháng {mm}-{yyyy}\\Danh sách+Chấm công\\[Chấm công {mm}.{yy_short}.xlsx]Công'!$C$5:$BC$6000,53,0),0)",
                "L": "=+IFERROR(VLOOKUP(C{row},'D:\\HỒ SƠ+LƯƠNG SD\\LƯƠNG\\LƯƠNG + CÔNG {yyyy}\\Tháng {mm}-{yyyy}\\Danh sách+Chấm công\\[Chấm công {mm}.{yy_short}.xlsx]Công'!$C$5:$BE$6000,55,0),0)",
                "M": "=+IFERROR(VLOOKUP(C{row},'D:\\HỒ SƠ+LƯƠNG SD\\LƯƠNG\\LƯƠNG + CÔNG {yyyy}\\Tháng {mm}-{yyyy}\\Danh sách+Chấm công\\[phụ cấp con nhỏ năm {yyyy}.xlsx]Con nhỏ'!$A$3:$B$3000,2,0),0)",
                "R": "=+IFERROR(VLOOKUP(C{row},'D:\\HỒ SƠ+LƯƠNG SD\\LƯƠNG\\LƯƠNG + CÔNG {yyyy}\\Tháng {mm}-{yyyy}\\Danh sách+Chấm công\\[phụ cấp con nhỏ năm {yyyy}.xlsx]Nhóm 6'!$C$2:$E$2000,3,0),0)",
                "T": "=+IFERROR(VLOOKUP(C{row},'D:\\HỒ SƠ+LƯƠNG SD\\LƯƠNG\\LƯƠNG + CÔNG {yyyy}\\Tháng {mm}-{yyyy}\\Danh sách+Chấm công\\[phụ cấp con nhỏ năm {yyyy}.xlsx]Thâm niên'!$B$3:$G$3000,6,0),0)",
                "U": "=+IFERROR(VLOOKUP(C{row},'D:\\HỒ SƠ+LƯƠNG SD\\LƯƠNG\\LƯƠNG + CÔNG {yyyy}\\Tháng {mm}-{yyyy}\\Danh sách+Chấm công\\[Chấm công {mm}.{yy_short}.xlsx]Công'!$C$5:$AP$6000,40,0),0)"
            }
        },
        "Lương {mm}": {
            "start_row": 10,
            "row_step": 2,
            "formulas": {
                "H": "=+(F{row}/26/8*{bep_cong_vu_rate})*H{row_minus_1}",
                "Y": "=+IFERROR(VLOOKUP(C{row_minus_1},'D:\\HỒ SƠ+LƯƠNG SD\\LƯƠNG\\LƯƠNG + CÔNG {prev_year}\\Tháng {prev_month}-{prev_year}\\[CK tháng {prev_month}.{prev_year}.xlsx]Bản gốc T{prev_month}'!$C$1:$J$1000,8,0),0)",
                "Z": "=+IFERROR(VLOOKUP(C{row_minus_1},'D:\\HỒ SƠ+LƯƠNG SD\\LƯƠNG\\LƯƠNG + CÔNG {prev_year}\\Tháng {prev_month}-{prev_year}\\[CK tháng {prev_month}.{prev_year}.xlsx]Bản gốc T{prev_month}'!$C$1:$F$1000,4,0),0)"
            }
        }
    }
}


# ==============================================================================
# Tự động điền các công thức động cấu hình theo template
# ==============================================================================
def apply_custom_formulas(wb, template_key, num_employees, mm, yyyy, to_name, log_callback):
    """
    Điền công thức động cho các sheet được cấu hình dựa trên template_key.
    """
    if template_key not in TEMPLATE_FORMULA_CONFIG:
        return

    log_callback(f"   ⚙️ Đang áp dụng công thức bổ sung cho template: {template_key}")
    
    # Tính toán các giá trị thời gian động
    yyyy_str = str(yyyy)
    mm_str = str(mm)
    mm_int = int(mm)
    
    # Tính tháng trước
    if mm_int == 1:
        prev_month_int = 12
        prev_year_int = int(yyyy) - 1
    else:
        prev_month_int = mm_int - 1
        prev_year_int = int(yyyy)
        
    prev_month = f"{prev_month_int:02d}"
    prev_year = str(prev_year_int)

    # Lấy 2 chữ số cuối của năm
    yy_short = yyyy_str[-2:]
    prev_yy_short = prev_year[-2:]

    # Xác định tỷ lệ động cho Bếp và Công vụ
    to_name_lower = str(to_name).lower()
    if "bếp" in to_name_lower:
        bep_cong_vu_rate = "125%"
    elif "công vụ" in to_name_lower:
        bep_cong_vu_rate = "115%"
    else:
        bep_cong_vu_rate = "100%"

    sheet_configs = TEMPLATE_FORMULA_CONFIG[template_key]
    
    excel_app = wb.Application
    link_cache = {}
    
    for sheet_name_pattern, s_config in sheet_configs.items():
        # Render tên sheet thực tế nếu có chứa {mm}
        resolved_sheet_name = sheet_name_pattern.replace("{mm}", mm_str)
        
        # Tìm sheet
        ws = None
        for sheet in wb.Sheets:
            if sheet.Name.lower() == resolved_sheet_name.lower():
                ws = sheet
                break
                
        if not ws:
            log_callback(f"   ⚠️ Không tìm thấy sheet '{resolved_sheet_name}' để áp dụng công thức.")
            continue
            
        start_row = s_config.get("start_row", 9)
        row_step = s_config.get("row_step", 2)
        formulas = s_config.get("formulas", {})
        
        # Mở khóa sheet nếu cần thiết
        try:
            ws.Unprotect()
        except Exception:
            pass
            
        # Điền công thức cho từng dòng nhân viên
        for idx in range(num_employees):
            current_row = start_row + (idx * row_step)
            
            for col_letter, formula_tmpl in formulas.items():
                formula = formula_tmpl.format(
                    row=current_row,
                    row_minus_1=current_row - 1,
                    row_plus_1=current_row + 1,
                    yyyy=yyyy_str,
                    yy_short=yy_short,
                    mm=mm_str,
                    mm_int=mm_int,
                    prev_year=prev_year,
                    prev_yy_short=prev_yy_short,
                    prev_month=prev_month,
                    prev_month_int=prev_month_int,
                    bep_cong_vu_rate=bep_cong_vu_rate
                )
                
                # Kiểm tra link ngoại (external links)
                link_match = re.search(r"'([^\[]*)\[([^\]]+)\]([^']+)'", formula)
                if link_match:
                    folder_path = link_match.group(1)
                    file_name = link_match.group(2)
                    sheet_name = link_match.group(3)
                    # Xử lý đường dẫn thực tế an toàn
                    actual_path = os.path.join(folder_path, file_name).replace("\\\\", "\\")
                    
                    is_valid = True
                    if actual_path not in link_cache:
                        if not os.path.exists(actual_path):
                            link_cache[actual_path] = None
                        else:
                            try:
                                temp_wb = excel_app.Workbooks.Open(actual_path, ReadOnly=True, UpdateLinks=0)
                                link_cache[actual_path] = [s.Name.lower() for s in temp_wb.Sheets]
                                temp_wb.Close(SaveChanges=False)
                            except Exception:
                                link_cache[actual_path] = None
                                
                    cached_sheets = link_cache[actual_path]
                    if cached_sheets is None:
                        if idx == 0:
                            log_callback(f"   ⚠️ Bỏ qua link: Không tìm thấy file '{actual_path}'. Để trống ô cột {col_letter}.")
                        is_valid = False
                    elif sheet_name.lower() not in cached_sheets:
                        if idx == 0:
                            log_callback(f"   ⚠️ Bỏ qua link: Không tìm thấy sheet '{sheet_name}' trong file '{file_name}'. Để trống ô cột {col_letter}.")
                        is_valid = False
                        
                    if not is_valid:
                        formula = ""
                
                # Ghi công thức bằng win32com
                cell_ref = f"{col_letter}{current_row}"
                try:
                    ws.Range(cell_ref).Value = formula
                except Exception as e:
                    log_callback(f"   ⚠️ Không thể ghi công thức tại {resolved_sheet_name}!{cell_ref}: {e}")
                    
        log_callback(f"   ✅ Đã tự động cập nhật công thức cho sheet '{resolved_sheet_name}' ({num_employees} dòng).")


# ==============================================================================
# Hàm chuyển đổi tên ô (VD: "O3") thành (row, col) số
# ==============================================================================
def cell_ref_to_rc(ref):
    """Chuyển đổi tham chiếu ô Excel (VD: 'O3') thành tuple (row, col)."""
    match = re.match(r'^([A-Z]+)(\d+)$', ref.upper())
    if not match:
        raise ValueError(f"Tham chiếu ô không hợp lệ: {ref}")
    col_str, row_str = match.groups()
    col = 0
    for ch in col_str:
        col = col * 26 + (ord(ch) - ord('A') + 1)
    return int(row_str), col


# ==============================================================================
# Tìm sheet "LƯƠNG Tx" (không phân biệt hoa thường) và đổi tên
# ==============================================================================
def find_and_rename_luong_sheet(wb, mm):
    """Tìm sheet có tên chứa 'lương t' (case-insensitive) và đổi thành 'Lương MM'."""
    target_name = f"Lương {mm}"
    for s in wb.Sheets:
        if s.Name.lower().startswith("lương t"):
            s.Name = target_name
            return s
    return None


# ==============================================================================
# Hàm xử lý Nhóm 3.1: to_may_template & kiem_hoa_template
# Tạo sheet "Danh sách", cập nhật Lương
# ==============================================================================
def process_group_31(wb, employees, mm, yyyy, config, log_callback):
    """Xử lý template Nhóm 3.1: cập nhật Lương + CĐ, rồi tạo sheet Danh sách ở cuối."""
    thang_nam = f"Tháng {mm} năm {yyyy}"

    # 1. Đổi tên sheet LƯƠNG Tx -> Lương MM (xử lý trước để không ảnh hưởng vị trí)
    ws_luong = find_and_rename_luong_sheet(wb, mm)
    if ws_luong:
        luong_ref = config.get("luong_cell", "O3")
        r, c = cell_ref_to_rc(luong_ref)
        ws_luong.Cells(r, c).Value = thang_nam
        log_callback(f"   ✅ Đã cập nhật sheet Lương: {luong_ref} = '{thang_nam}'")
    else:
        log_callback("   ⚠️ Không tìm thấy sheet 'LƯƠNG Tx'")

    # 2. Tạo sheet "Danh sách" ở vị trí CUỐI CÙNG (sau tất cả sheet khác)
    try:
        # Truyền theo vị trí để tránh lỗi late-binding của win32com: Add(Before, After) -> (None, last_sheet)
        last_sheet = wb.Sheets(wb.Sheets.Count)
        ws_ds = wb.Worksheets.Add(None, last_sheet)
        ws_ds.Name = "Danh sách"
        log_callback(f"   📋 Đã tạo sheet 'Danh sách' ở vị trí cuối (sheet thứ {wb.Sheets.Count})")

        # Ghi header
        headers = ["STT", "HN/TTN", "Tổ/bộ phận", "Mã số NV", "Họ và Tên"]
        for i, h in enumerate(headers):
            cell = ws_ds.Cells(1, i + 1)
            cell.Value = h
            cell.Font.Bold = True

        # Ghi dữ liệu nhân viên
        for idx, emp in enumerate(employees):
            row = idx + 2
            ws_ds.Cells(row, 1).Value = idx + 1         # STT
            ws_ds.Cells(row, 2).Value = emp["hn_ttn"]   # HN/TTN (Cột B)
            ws_ds.Cells(row, 3).Value = emp["to_bp"]    # Tổ/bộ phận (Cột C)
            ws_ds.Cells(row, 4).Value = emp["ma_nv"]    # Mã số NV (Cột D)
            ws_ds.Cells(row, 5).Value = emp["ho_ten"]   # Họ và Tên (Cột E)

        ws_ds.Columns("A:E").AutoFit()
    except Exception as e:
        log_callback(f"   ⚠️ Lỗi tạo sheet Danh sách: {e}")


# ==============================================================================
# Hàm xử lý Nhóm 3.2 -> 3.7: Các template GT
# Điền nhân viên vào "Bang TH nop" (cách 1 dòng trống), cập nhật Lương + CĐ
# ==============================================================================
def process_group_gt(wb, employees, mm, yyyy, to_name, config, log_callback):
    """Xử lý template GT: điền Bang TH nop + cập nhật Lương + CĐ."""
    thang_nam = f"Tháng {mm} năm {yyyy}"

    # 1. Thao tác trên sheet "Bang TH nop"
    try:
        ws = wb.Sheets("Bang TH nop")

        # Cập nhật tiêu đề A2 nếu cần
        if config.get("update_a2", False):
            a2_text = f"BẢNG HỆ SỐ HƯỞNG LƯƠNG THÁNG {mm}/{yyyy}"
            ws.Cells(2, 1).Value = a2_text
            log_callback(f"   ✅ A2 = '{a2_text}'")

        # Cập nhật A3 nếu cần
        if config.get("update_a3", False):
            if "a3_fixed" in config:
                a3_text = config["a3_fixed"]
            else:
                a3_text = f"Tổ {to_name}"
            ws.Cells(3, 1).Value = a3_text
            log_callback(f"   ✅ A3 = '{a3_text}'")

        # Điền danh sách nhân viên từ dòng 9, cách nhau 1 dòng trống (ROW_STEP=2)
        START_ROW = 9
        ROW_STEP = 2
        for idx, emp in enumerate(employees):
            current_row = START_ROW + (idx * ROW_STEP)

            # Đảm bảo dòng không bị ẩn
            ws.Rows(current_row).Hidden = False
            ws.Rows(current_row).RowHeight = 15

            ws.Cells(current_row, 1).Value = idx + 1        # STT
            ws.Cells(current_row, 2).Value = emp["ho_ten"]  # Họ và Tên
            ws.Cells(current_row, 3).Value = emp["ma_nv"]   # Mã số NV

        log_callback(f"   ✅ Đã điền {len(employees)} nhân viên vào 'Bang TH nop' (dòng {START_ROW} -> {START_ROW + (len(employees)-1)*ROW_STEP})")
    except Exception as e:
        log_callback(f"   ⚠️ Lỗi xử lý Bang TH nop: {e}")

    # 2. Đổi tên sheet LƯƠNG Tx -> Lương MM và cập nhật ô tương ứng
    ws_luong = find_and_rename_luong_sheet(wb, mm)
    if ws_luong:
        luong_ref = config.get("luong_cell", "N3")
        r, c = cell_ref_to_rc(luong_ref)
        ws_luong.Cells(r, c).Value = thang_nam
        log_callback(f"   ✅ Đã cập nhật sheet Lương: {luong_ref} = '{thang_nam}'")
    else:
        log_callback("   ⚠️ Không tìm thấy sheet 'LƯƠNG Tx'")


# ==============================================================================
# Hàm xử lý template "chua_xu_ly_template.xlsx" (trường hợp còn lại)
# ==============================================================================
def process_chua_xu_ly(wb, employees, mm, yyyy, log_callback):
    """Ghi danh sách nhân viên chưa được xử lý vào template."""
    try:
        ws = wb.Sheets(1)  # Lấy sheet đầu tiên
        # Ghi dữ liệu từ dòng 2 (dòng 1 là header)
        for idx, emp in enumerate(employees):
            row = idx + 2
            ws.Cells(row, 1).Value = idx + 1
            ws.Cells(row, 2).Value = emp["hn_ttn"]
            ws.Cells(row, 3).Value = emp["to_bp"]
            ws.Cells(row, 4).Value = emp["ma_nv"]
            ws.Cells(row, 5).Value = emp["ho_ten"]
        log_callback(f"   ✅ Đã ghi {len(employees)} nhân viên chưa xử lý")
    except Exception as e:
        log_callback(f"   ⚠️ Lỗi ghi dữ liệu chưa xử lý: {e}")


# ==============================================================================
# Đảm bảo định dạng cho các ô tiêu đề đã merge (R4:R7, S4:S7, T4:T7) trong sheet 'Tong hop cac ma'
# ==============================================================================
def ensure_merged_headers_format(wb):
    """
    Đảm bảo định dạng cho các ô tiêu đề đã merge (R4:R7, S4:S7, T4:T7) trong sheet 'Tong hop cac ma'.
    Excel COM có thể bị mất định dạng của ô góc trên bên trái (R4, S4, T4) sau khi lưu nếu không được gán tường minh.
    """
    try:
        ws_tong_hop = None
        for s in wb.Sheets:
            if s.Name == "Tong hop cac ma":
                ws_tong_hop = s
                break
        
        if ws_tong_hop:
            # Unprotect sheet trước khi định dạng
            try: ws_tong_hop.Unprotect()
            except: pass
            
            cells_to_format = ["R4", "S4", "T4"]
            for cell_ref in cells_to_format:
                cell = ws_tong_hop.Range(cell_ref)
                # Căn lề giữa
                cell.HorizontalAlignment = -4108 # xlCenter
                cell.VerticalAlignment = -4108   # xlCenter
                # Chữ in đậm
                cell.Font.Bold = True
                # Tự động xuống dòng
                cell.WrapText = True
                # Màu nền vàng (RGB: 255, 255, 0 = 65535)
                cell.Interior.Color = 65535
                # Màu chữ đỏ (RGB: 255, 0, 0 = 255)
                cell.Font.Color = 255
    except Exception:
        pass


import win32gui
import win32process
import win32com.client as win32

def kill_background_excel(log_callback):
    """Tìm và kill các tiến trình EXCEL.EXE chạy ngầm (không có cửa sổ hiển thị) do các lần chạy trước bị crash để lại."""
    try:
        excel_pids_with_windows = set()
        
        def enum_windows_proc(hwnd, lParam):
            try:
                if win32gui.IsWindowVisible(hwnd):
                    if win32gui.GetClassName(hwnd) == "XLMAIN":
                        _, pid = win32process.GetWindowThreadProcessId(hwnd)
                        excel_pids_with_windows.add(pid)
            except Exception:
                pass
                    
        win32gui.EnumWindows(enum_windows_proc, 0)
        
        # Dùng WMI (có sẵn trên Windows) để lấy các PID của EXCEL.EXE
        wmi = win32.GetObject('winmgmts:')
        processes = wmi.ExecQuery("Select ProcessId from Win32_Process where Name='EXCEL.EXE'")
        
        killed_count = 0
        for p in processes:
            pid = int(p.ProcessId)
            if pid not in excel_pids_with_windows:
                # Tiêu diệt tiến trình ngầm bằng CMD
                os.system(f"taskkill /f /pid {pid} >nul 2>&1")
                killed_count += 1
                
        if killed_count > 0:
            log_callback(f"   💀 Đã dọn dẹp {killed_count} tiến trình Excel chạy ngầm bị treo từ trước.")
    except Exception as e:
        log_callback(f"   ⚠️ Lỗi khi dọn dẹp Excel ngầm: {e}")

# ==============================================================================
# Hàm chính: Ghi dữ liệu đã phân loại vào các file Excel dựa trên Template
# ==============================================================================
def close_open_excel_files_in_dir(target_dir, log_callback):
    """Đóng các file Excel đang mở thuộc thư mục target_dir để tránh lỗi file đang được sử dụng."""
    try:
        excel = win32.GetActiveObject("Excel.Application")
        target_dir_lower = os.path.abspath(target_dir).lower()
        closed = 0
        
        wbs_to_close = []
        for i in range(1, excel.Workbooks.Count + 1):
            try:
                wb = excel.Workbooks.Item(i)
                if str(wb.FullName).lower().startswith(target_dir_lower):
                    wbs_to_close.append(wb)
            except Exception:
                pass
                
        for wb in wbs_to_close:
            try:
                filename = wb.Name
                wb.Close(SaveChanges=False)
                closed += 1
                log_callback(f"   🧹 Đã đóng file đang mở: {filename}")
            except Exception:
                pass
                
        if closed > 0:
            log_callback(f"   ✅ Đã dọn dẹp {closed} file cũ đang mở trong thư mục.")
    except Exception:
        pass


def write_all_groups(classified_data, output_dir, mm, yyyy, log_callback, progress_callback=None):
    """
    Duyệt qua từng nhóm đã phân loại, copy template, ghi dữ liệu và lưu file.
    classified_data: dict từ classify_employees()
    """
    total = len(classified_data)
    if total == 0:
        return True, "Không có dữ liệu để xử lý."

    pythoncom.CoInitialize()
    
    # Dọn dẹp các tiến trình Excel ngầm bị treo từ lần chạy trước
    kill_background_excel(log_callback)
    
    # Tự động đóng các file trong thư mục output đang bị mở (ở Excel hiển thị)
    close_open_excel_files_in_dir(output_dir, log_callback)
    
    processed = 0
    excel_app = None

    try:
        # Khởi tạo ứng dụng Excel ẩn
        excel_app = win32.DispatchEx('Excel.Application')
        excel_app.Visible = False
        excel_app.DisplayAlerts = False
        excel_app.ScreenUpdating = False
        excel_app.AskToUpdateLinks = False
        excel_app.Interactive = False

        for filename, group_info in classified_data.items():
            template_key = group_info["template_key"]
            template_path = group_info["template_path"]
            employees = group_info["employees"]
            to_name = group_info["to_name"]

            log_callback(f"\n--- Đang xử lý: {filename} ({len(employees)} NV) ---")

            # Kiểm tra template tồn tại
            if not os.path.exists(template_path):
                log_callback(f"   ❌ Không tìm thấy template: {template_path}")
                continue

            # Tạo đường dẫn file output chuẩn xác theo phân loại (sx/, gt/ hoặc thư mục gốc)
            sub_folder = None
            if template_key.startswith("sx/"):
                sub_folder = "sx"
            elif template_key.startswith("gt/"):
                sub_folder = "gt"
            
            if sub_folder:
                target_folder = os.path.join(output_dir, sub_folder)
            else:
                target_folder = output_dir  # Danh sách chưa được xử lý lưu trực tiếp ở thư mục gốc
                
            os.makedirs(target_folder, exist_ok=True)
            output_path = os.path.join(target_folder, filename)

            # Xóa file cũ nếu tồn tại
            if os.path.exists(output_path):
                try:
                    os.remove(output_path)
                except PermissionError:
                    log_callback(f"   ❌ File đang mở, không thể ghi đè: {filename}")
                    continue

            # Copy template sang output
            shutil.copy2(template_path, output_path)

            # Mở file bằng win32com
            wb = None
            try:
                wb = excel_app.Workbooks.Open(os.path.abspath(output_path), UpdateLinks=0)
                config = TEMPLATE_CONFIG.get(template_key, {})

                if template_key == "chua_xu_ly_template.xlsx":
                    # Trường hợp chưa xử lý
                    process_chua_xu_ly(wb, employees, mm, yyyy, log_callback)

                elif config.get("has_danh_sach", False):
                    # Nhóm 3.1: to_may_template & kiem_hoa_template
                    process_group_31(wb, employees, mm, yyyy, config, log_callback)

                else:
                    # Nhóm 3.2 -> 3.7: Các template GT
                    process_group_gt(wb, employees, mm, yyyy, to_name, config, log_callback)

                # Áp dụng công thức đặc biệt cấu hình động cho các template (nếu có)
                apply_custom_formulas(wb, template_key, len(employees), mm, yyyy, to_name, log_callback)

                # Đảm bảo định dạng cho các ô tiêu đề đã merge trước khi Save
                ensure_merged_headers_format(wb)
                
                wb.Save()
                log_callback(f"   ✅ Đã lưu: {filename}")
                processed += 1

            except Exception as e:
                log_callback(f"   ❌ Lỗi xử lý {filename}: {e}")
            finally:
                if wb is not None:
                    try:
                        wb.Close(SaveChanges=False)
                    except Exception:
                        pass
                    del wb

            # Cập nhật tiến độ
            if progress_callback:
                progress_callback(int(((processed) / total) * 100))

        if progress_callback:
            progress_callback(100)

    finally:
        if excel_app is not None:
            try:
                excel_app.Quit()
            except Exception:
                pass
            del excel_app
        pythoncom.CoUninitialize()

    return True, f"Hoàn tất! Đã tạo {processed}/{total} file thành công."
