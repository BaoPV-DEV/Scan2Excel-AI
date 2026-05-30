import os
import shutil
import re
import json
import pythoncom
import win32com.client as win32
import win32gui
import win32process


# ==============================================================================
# TEMPLATE CONFIG
# ==============================================================================
TEMPLATE_CONFIG = {
    "sx/to_may_template.xlsx": {"has_danh_sach": True, "luong_cell": "O3"},
    "sx/kiem_hoa_template.xlsx": {"has_danh_sach": True, "luong_cell": "O3"},

    "gt/bao_ve_template.xlsx": {"has_danh_sach": False, "luong_cell": "O3"},
    "gt/bep_an_va_cong_vu_template.xlsx": {"has_danh_sach": False, "luong_cell": "K3"},
    "gt/boc_vac_template.xlsx": {"has_danh_sach": False, "luong_cell": "N4", "a3_fixed": "Tổ Bốc vác"},
    "gt/co_dien_template.xlsx": {"has_danh_sach": False, "luong_cell": "N4", "a3_fixed": "Tổ Cơ điện"},
    "gt/ke_hoach_template.xlsx": {"has_danh_sach": False, "luong_cell": "O4"},
    "gt/ky_thuat_template.xlsx": {"has_danh_sach": False, "luong_cell": "O4"},
    "gt/van_phong_template.xlsx": {"has_danh_sach": False, "luong_cell": "O4", "a3_fixed": "Bộ phận Văn phòng"},
}


# ==============================================================================
# UTILS
# ==============================================================================
def safe_str(v):
    if v is None:
        return ""
    try:
        if str(v).lower() in ["nan", "<na>", "none"]:
            return ""
        return str(v).strip()
    except:
        return ""


def batch_write(ws, start_cell, data):
    if not data:
        return

    if isinstance(data[0], (str, int, float)):
        data = [data]

    ws.Range(start_cell).Resize(len(data), len(data[0])).Value = data


def cell_ref_to_rc(ref):
    m = re.match(r'^([A-Z]+)(\d+)$', ref.upper())
    if not m:
        raise ValueError(f"Tham chiếu ô không hợp lệ: {ref}")

    col_str, row_str = m.groups()
    col = 0
    for c in col_str:
        col = col * 26 + (ord(c) - 64)

    return int(row_str), col


# ==============================================================================
# EXCEL CLEANER
# ==============================================================================
def kill_excel_safe(log):
    try:
        visible_pids = set()

        def enum(hwnd, _):
            if win32gui.IsWindowVisible(hwnd):
                if win32gui.GetClassName(hwnd) == "XLMAIN":
                    _, pid = win32process.GetWindowThreadProcessId(hwnd)
                    visible_pids.add(pid)

        win32gui.EnumWindows(enum, 0)

        wmi = win32.GetObject("winmgmts:")
        procs = wmi.ExecQuery("Select ProcessId from Win32_Process where Name='EXCEL.EXE'")

        killed = 0
        for p in procs:
            pid = int(p.ProcessId)
            if pid not in visible_pids:
                os.system(f"taskkill /f /pid {pid} >nul 2>&1")
                killed += 1

        if killed > 0:
            log(f"💀 Đã dọn {killed} tiến trình Excel bị treo từ lần chạy trước.")

    except Exception as e:
        log(f"⚠️ Lỗi khi dọn Excel ngầm: {e}")


# ==============================================================================
# GROUP 3.1 (SX - có sheet Danh sách)
# ==============================================================================
def process_group_31(wb, employees, mm, yyyy, config, log):
    log("📄 Tạo sheet 'Danh sách'...")

    # 1. Tạo và đưa sheet xuống cuối cùng
    ws = create_sheet_last(wb, "Danh sách")
    ws.Activate()  # Ép Excel tập trung vào sheet này để ghi dữ liệu công khai

    # 2. Ghi tiêu đề (Header phải là mảng 2 chiều: 1 dòng, 5 cột)
    headers = (("STT", "HN/TTN", "Tổ/bộ phận", "Mã số NV", "Họ và Tên"),)
    ws.Range("A1:E1").Value = headers

    log(f"   ➜ Chuẩn bị ghi {len(employees)} nhân viên...")

    # 3. Tạo mảng dữ liệu 2 chiều (Mỗi nhân viên là một tuple bên trong)
    data_matrix = []
    for i, e in enumerate(employees):
        data_matrix.append((
            int(i + 1),                           # Cột STT (Kiểu số)
            str(e.get("hn_ttn", "")).strip(),     # Cột HN/TTN
            str(e.get("to_bp", "")).strip(),      # Cột Tổ/bộ phận
            str(e.get("ma_nv", "")).strip(),      # Cột Mã số NV
            str(e.get("ho_ten", "")).strip()      # Cột Họ và Tên
        ))

    if data_matrix:
        # CHỐNG LỖI LỆCH Ô: Tính toán chính xác dòng bắt đầu và dòng kết thúc
        start_row = 2
        end_row = start_row + len(data_matrix) - 1
        
        # Tạo chuỗi định vị dải ô chính xác, ví dụ: "A2:E18"
        target_range_str = f"A{start_row}:E{end_row}"
        log(f"   ➜ Đang đẩy dữ liệu trực tiếp vào vùng: {target_range_str}")
        
        # Ghi đè trực tiếp toàn bộ mảng 2 chiều vào đúng dải ô quy định
        ws.Range(target_range_str).Value = tuple(data_matrix)

    # 4. Tự động căn chỉnh độ rộng cột và bật làm tươi màn hình Excel
    ws.Columns("A:E").AutoFit()
    wb.Application.ScreenUpdating = True
    
    log("   ✅ Hoàn tất ghi dữ liệu vào sheet 'Danh sách'")


def create_sheet_last(wb, sheet_name="Danh sách"):
    # 1. Tắt cảnh báo để xóa sheet không bị hiện thông báo Excel
    old_alerts = wb.Application.DisplayAlerts
    wb.Application.DisplayAlerts = False
    
    try:
        for s in wb.Sheets:
            if s.Name == sheet_name:
                s.Delete()
                break
    finally:
        wb.Application.DisplayAlerts = old_alerts

    # 2. Đặt tham số rõ ràng cho hàm Add: (Before, After) -> Không dùng Before nên truyền None hoặc để trống
    # Trong win32com, để truyền tham số thứ hai (After), ta truyền đối tượng sheet cuối vào tham số thứ hai.
    last_sheet = wb.Sheets(wb.Sheets.Count)
    ws = wb.Worksheets.Add(None, last_sheet) 

    ws.Name = sheet_name
    return ws


# ==============================================================================
# GROUP GT
# ==============================================================================
def process_group_gt(wb, employees, mm, yyyy, to_name, config, log):
    log(f"📄 Đang xử lý nhóm GT: {to_name}")

    try:
        ws = wb.Sheets("Bang TH nop")
    except:
        log("❌ Không tìm thấy sheet 'Bang TH nop'")
        return

    if config.get("a3_fixed"):
        ws.Cells(3, 1).Value = config["a3_fixed"]
        log(f"   ➜ A3 cố định: {config['a3_fixed']}")
    else:
        ws.Cells(3, 1).Value = f"Tổ {to_name}"
        log(f"   ➜ A3: Tổ {to_name}")

    START = 9
    STEP = 2

    log(f"   ➜ Bắt đầu ghi nhân viên từ dòng {START} (cách {STEP} dòng)")

    for i, e in enumerate(employees):
        r = START + i * STEP
        ws.Cells(r, 1).Value = i + 1
        ws.Cells(r, 2).Value = safe_str(e.get("ho_ten"))
        ws.Cells(r, 3).Value = safe_str(e.get("ma_nv"))

    log(f"   ✅ Đã ghi {len(employees)} nhân viên vào 'Bang TH nop'.")


# ==============================================================================
# MAIN PIPELINE
# ==============================================================================
def write_all_groups(classified_data, output_dir, mm, yyyy, log, progress=None):
    pythoncom.CoInitialize()

    log("🚀 Bắt đầu xử lý xuất file Excel...")
    kill_excel_safe(log)

    excel = win32.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    
    # Ép Excel bỏ qua chế độ xem được bảo vệ (Tránh lỗi đơ khi mở file tải từ mạng/mở đồng thời)
    excel.AutomationSecurity = 1 # 1 = msoAutomationSecurityLow

    total = len(classified_data)
    done = 0

    log(f"📦 Tổng số file cần tạo: {total}")

    for filename, g in classified_data.items():
        wb = None
        try:
            template = g["template_path"]
            employees = g["employees"]
            to_name = g["to_name"]
            key = g["template_key"]

            log("\n" + "=" * 60)
            log(f"📄 Đang xử lý file: {filename}")
            log(f"   ➜ Số nhân viên: {len(employees)}")
            log(f"   ➜ Template: {template}")

            if not os.path.exists(template):
                log("❌ Không tìm thấy file template! Bỏ qua file này.")
                continue

            folder = os.path.join(output_dir, "sx" if key.startswith("sx") else "gt")
            os.makedirs(folder, exist_ok=True)

            out = os.path.join(folder, filename)

            # Khối try xóa file cũ, nếu file đang mở, hệ thống sẽ báo lỗi và nhảy sang file tiếp theo
            if os.path.exists(out):
                try:
                    os.remove(out)
                    log("   🧹 Đã xóa file cũ")
                except PermissionError:
                    log(f"❌ File '{filename}' đang mở ở một chương trình khác và bị khóa. BỎ QUA FILE NÀY.")
                    continue

            shutil.copy2(template, out)
            log("   📌 Đã copy template sang file output")

            # Mở file bằng Excel
            wb = excel.Workbooks.Open(out, UpdateLinks=0)

            cfg = TEMPLATE_CONFIG.get(key, {})

            if cfg.get("has_danh_sach"):
                process_group_31(wb, employees, mm, yyyy, cfg, log)
            else:
                process_group_gt(wb, employees, mm, yyyy, to_name, cfg, log)

            wb.Save()
            done += 1
            log(f"✔ Lưu thành công: {filename}")

        except Exception as file_error:
            # CHỐNG CRASH TOÀN BỘ APP: Ghi nhận lỗi của riêng file này và tiếp tục vòng lặp
            log(f"💥 LỖI KHI XỬ LÝ FILE '{filename}': {file_error}")
            log("   ➜ Bỏ qua file này và tiếp tục xử lý các file khác...")
            
        finally:
            # Đảm bảo đóng workbook hiện tại dù thành công hay thất bại để giải phóng Excel
            if wb:
                try:
                    wb.Close(False)
                except:
                    pass

        if progress:
            progress(int(done / total * 100))

    if progress:
        progress(100)

    log("\n🎉 HOÀN TẤT TOÀN BỘ QUÁ TRÌNH XỬ LÝ FILE EXCEL")
    log(f"✔ Thành công: {done}/{total} file")

    try:
        excel.Quit()
    except:
        pass
        
    pythoncom.CoUninitialize()
    return True, f"Hoàn tất {done}/{total} file"