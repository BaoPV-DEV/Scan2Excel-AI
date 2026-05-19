import os
import shutil
import re
import pythoncom
import win32com.client as win32
from logic.tab1_classifier import TEMPLATE_CONFIG

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


# ==============================================================================
# Hàm chính: Ghi dữ liệu đã phân loại vào các file Excel dựa trên Template
# ==============================================================================
def write_all_groups(classified_data, output_dir, mm, yyyy, log_callback, progress_callback=None):
    """
    Duyệt qua từng nhóm đã phân loại, copy template, ghi dữ liệu và lưu file.
    classified_data: dict từ classify_employees()
    """
    total = len(classified_data)
    if total == 0:
        return True, "Không có dữ liệu để xử lý."

    processed = 0
    pythoncom.CoInitialize()
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
