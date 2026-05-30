import gc
import json
import os
import re
import pythoncom
import win32com.client as win32
from dotenv import load_dotenv
from functools import lru_cache

# =========================================================
# 1. NORMALIZE (CACHE để giảm CPU)
# =========================================================
@lru_cache(maxsize=5000)
def normalize_str(s):
    """Chuẩn hóa chuỗi (bỏ dấu + lowercase + trim)"""
    if not s:
        return ""
    s = str(s).strip().lower()

    s = re.sub(r'[àáạảãâầấậẩẫăằắặẳẵ]', 'a', s)
    s = re.sub(r'[èéẹẻẽêềếệểễ]', 'e', s)
    s = re.sub(r'[ìíịỉĩ]', 'i', s)
    s = re.sub(r'[òóọỏõôồốộổỗơờớợởỡ]', 'o', s)
    s = re.sub(r'[ùúụủũưừứựửữ]', 'u', s)
    s = re.sub(r'[ỳýỵỷỹ]', 'y', s)
    s = re.sub(r'[đ]', 'd', s)
    s = re.sub(r'\s+', ' ', s)

    return s

# =========================================================
# 2. MATCH TEAM LOGIC
# =========================================================
def teams_match(norm_json, norm_header):
    """So sánh tên tổ đã normalize"""
    if norm_json == norm_header:
        return True

    if norm_header in ["co dong", "dieu dong"] and norm_json in ["co dong", "dieu dong"]:
        return True

    if norm_header in ["hoan thien", "h.thien"] and norm_json in ["hoan thien", "h.thien"]:
        return True

    return False

# =========================================================
# 3. FORMAT HEADER MERGED CELLS
# =========================================================
def ensure_merged_headers_format(wb):
    """Fix format cho header merge trong Excel"""
    try:
        ws = None
        for s in wb.Sheets:
            if s.Name == "Tong hop cac ma":
                ws = s
                break

        if ws:
            try:
                ws.Unprotect()
            except:
                pass

            for cell in ["R4", "S4", "T4"]:
                c = ws.Range(cell)
                c.HorizontalAlignment = -4108  # xlCenter
                c.VerticalAlignment = -4108    # xlCenter
                c.Font.Bold = True
                c.WrapText = True
                c.Interior.Color = 65535       # Yellow
                c.Font.Color = 255             # Red
    except:
        pass

# =========================================================
# 4. MAIN ENGINE (WITH INTENSE DEBUG LOGS)
# =========================================================
def process_excel_integration(json_folder, excel_target_folder, log_callback,
                             progress_callback=None, excel_pass="", single_dept=None):

    load_dotenv()
    EXCEL_PASS = excel_pass or os.getenv("EXCEL_SHEET_PASSWORD", "8863")

    log_callback(f"[DEBUG START] Gốc JSON: {json_folder}")
    log_callback(f"[DEBUG START] Thư mục Excel: {excel_target_folder}")
    log_callback(f"[DEBUG START] Bộ lọc đơn tổ (single_dept): {single_dept}")

    norm_single_dept = normalize_str(single_dept) if single_dept else None

    # ================================
    # 1. LOAD JSON FILES
    # ================================
    json_paths = []
    for root, _, files in os.walk(json_folder):
        for f in files:
            if f.endswith(".json"):
                json_paths.append(os.path.join(root, f))

    log_callback(f"[DEBUG] Tổng số file JSON tìm thấy trên ổ đĩa: {len(json_paths)}")
    if not json_paths:
        return False, "Không tìm thấy JSON"

    excel_groups = {}
    employee_dept_giay = {}

    # ================================
    # 2. GROUP JSON → EXCEL & TÍNH TỔNG GIÂY
    # ================================
    for jp in json_paths:
        try:
            with open(jp, "r", encoding="utf-8") as f:
                data = json.load(f)

            info = data.get("thong_tin_chung", {})
            to_sx = str(info.get("to_san_xuat", "")).strip()
            if to_sx.isdigit():
                to_sx = f"Tổ {to_sx}"

            to_sx_upper = to_sx.upper()
            norm_to_sx = normalize_str(to_sx_upper)

            # Lọc đơn tổ (nếu có)
            if norm_single_dept and not teams_match(norm_to_sx, norm_single_dept):
                continue

            thuong = float(info.get("thuong_ma_hang_moi", 0)) / 100.0
            
            # Đếm log số công đoạn
            cong_doan_list = data.get("danh_sach_cong_doan", [])

            for cd in cong_doan_list:
                dinh_muc = float(cd.get("dinh_muc_t", 0))
                if dinh_muc <= 0:
                    continue

                for th in cd.get("thuc_hien", []):
                    m_id = str(th.get("ma_nhan_vien", "")).strip().replace(".0", "")
                    sl = float(th.get("so_luong", 0))

                    if m_id and sl > 0:
                        giay = (sl * dinh_muc / 100.0) * (1 + thuong)

                        if m_id not in employee_dept_giay:
                            employee_dept_giay[m_id] = {}

                        employee_dept_giay[m_id][to_sx_upper] = \
                            employee_dept_giay[m_id].get(to_sx_upper, 0) + giay

            # Định vị file Excel
            match = re.search(r"(\d{2})/(\d{4})", str(info.get("thoi_gian", "")))
            if not match:
                log_callback(f"[WARNING] File JSON {os.path.basename(jp)} sai định dạng 'thoi_gian'")
                continue

            mm = match.group(1)
            excel_name = f"{to_sx} - {mm}.xlsx"
            target_path = os.path.abspath(os.path.join(excel_target_folder, excel_name))

            excel_groups.setdefault(target_path, []).append({
                "path": jp,
                "data": data,
                "to_sx_upper": to_sx_upper
            })

        except Exception as e:
            log_callback(f"[ERROR] Lỗi phân tích JSON {os.path.basename(jp)}: {e}")
            continue

    log_callback(f"[DEBUG] Số nhóm file Excel đích sẽ xử lý: {len(excel_groups)}")
    for k, v in excel_groups.items():
        log_callback(f"   -> Excel: {os.path.basename(k)} (Gồm {len(v)} item hàng)")

    if not excel_groups:
        return False, "Không có dữ liệu phù hợp để tích hợp (Có thể do bộ lọc Tổ trống)"

    # ================================
    # 3. INIT EXCEL COM
    # ================================
    pythoncom.CoInitialize()
    excel_app = None

    processed_count = 0
    total_work = sum(len(v) for v in excel_groups.values())

    try:
        excel_app = win32.DispatchEx("Excel.Application")
        excel_app.Visible = False
        excel_app.DisplayAlerts = False
        excel_app.ScreenUpdating = False

        all_paths = list(excel_groups.keys())

        # ================================
        # 4. PROCESS EACH EXCEL FILE
        # ================================
        for excel_path in all_paths:
            log_callback(f"\n[EXCEL] Bắt đầu mở file: {os.path.basename(excel_path)}")
            if not os.path.exists(excel_path):
                log_callback(f"[ERROR] KHÔNG TÌM THẤY FILE TRÊN ĐĨA: {excel_path}")
                continue

            wb = None
            try:
                wb = excel_app.Workbooks.Open(excel_path)
                ensure_merged_headers_format(wb)

                try:
                    ws = wb.Sheets("Bang TH nop")
                except Exception as sheet_err:
                    log_callback(f"[ERROR] Lỗi mở Sheet 'Bang TH nop': {sheet_err}")
                    continue

                try:
                    ws.Unprotect(EXCEL_PASS)
                    log_callback(f"[DEBUG] Đã Unprotect sheet bằng pass: {EXCEL_PASS}")
                except Exception as unprotect_err:
                    log_callback(f"[WARNING] Không thể Unprotect sheet (Có thể sheet không khóa): {unprotect_err}")

                # Tìm hàng cuối cùng chứa dữ liệu ở cột B
                last_row = ws.Cells(ws.Rows.Count, "B").End(-4162).Row  # -4162 = xlUp
                log_callback(f"[DEBUG] Dòng cuối tìm thấy tại cột B: Dòng {last_row}")
                if last_row < 7: 
                    log_callback("[WARNING] Dòng cuối < 7, tự động đặt giả định bằng 200 dòng.")
                    last_row = 200

                # Đọc danh sách định vị
                id_range_values = ws.Range(f"B7:B{last_row}").Value
                header_range_values = ws.Range("C5:Z5").Value

                # Ánh xạ Mã NV -> Chỉ mục
                row_map = {}
                if id_range_values:
                    for idx, val in enumerate(id_range_values):
                        if val is not None:
                            clean_id = str(val).strip().replace(".0", "")
                            row_map[clean_id] = idx
                log_callback(f"[DEBUG] Đã nạp thành công {len(row_map)} Mã nhân viên từ cột B của Excel vào bộ nhớ.")

                # Ánh xạ Tên Tổ -> Chỉ mục
                col_map = {}
                if header_range_values and header_range_values[0]:
                    for idx, val in enumerate(header_range_values[0]):
                        if val is not None:
                            norm_header = normalize_str(val)
                            col_map[norm_header] = idx
                            log_callback(f"   -> Quét thấy cột Tổ Excel: '{val}' (Normalize: '{norm_header}') -> Chỉ mục mảng: {idx}")

                # Tải khối ma trận dữ liệu từ C7 đến Z[last_row]
                data_range = ws.Range(f"C7:Z{last_row}")
                current_values = data_range.Value
                grid_data = [list(row) for row in current_values] if current_values else []

                if not grid_data:
                    log_callback(f"[ERROR] Mảng dữ liệu lưới ô rỗng, không thể chỉnh sửa.")
                    continue

                # ================================
                # 5. ĐIỀN DỮ LIỆU VÀO MẢNG TRÊN RAM
                # ================================
                items = excel_groups.get(excel_path, [])
                for item in items:
                    data = item["data"]
                    info = data["thong_tin_chung"]
                    ma_hang = info.get("ma_hang", "UNKNOWN")

                    log_callback(f"[MÃ HÀNG] Đang xử lý: {ma_hang}")

                    for cd in data.get("danh_sach_cong_doan", []):
                        for th in cd.get("thuc_hien", []):
                            m_id = str(th.get("ma_nhan_vien", "")).strip().replace(".0", "")
                            sl = float(th.get("so_luong", 0))

                            if not m_id or sl <= 0:
                                continue

                            # Kiểm tra sự tồn tại của Mã NV trong Excel
                            if m_id in row_map:
                                array_row_idx = row_map[m_id]
                                dept_giay_dict = employee_dept_giay.get(m_id, {})

                                for json_dept, total_giay in dept_giay_dict.items():
                                    norm_json_dept = normalize_str(json_dept)
                                    is_mapped_col = False

                                    for norm_header, array_col_idx in col_map.items():
                                        if teams_match(norm_json_dept, norm_header):
                                            
                                            # LOG ĐỐI CHIẾU GIÁ TRỊ TRƯỚC/SAU KHI GHI
                                            val_cu = grid_data[array_row_idx][array_col_idx]
                                            grid_data[array_row_idx][array_col_idx] = total_giay
                                            is_mapped_col = True
                                            
                                            log_callback(f"   [OK MATCH] Nhân viên {m_id} | Tổ JSON: '{json_dept}' khớp với cột Excel: '{norm_header}' -> Ghi giá trị: {total_giay} (Cũ: {val_cu})")
                                            break
                                    
                                    if not is_mapped_col:
                                        log_callback(f"   [FAIL COL] Nhân viên {m_id} có Tổ JSON '{json_dept}' nhưng KHÔNG KHỚP với bất kỳ cột nào trên dòng 5 Excel!")
                            else:
                                # Log cảnh báo nếu mã nhân viên trong JSON không có trong file Excel
                                log_callback(f"   [FAIL ROW] Mã NV '{m_id}' trong JSON không tìm thấy tại cột B (Từ dòng 7) của file Excel này!")

                    processed_count += 1
                    if progress_callback and total_work:
                        progress_callback(int(processed_count * 100 / total_work))

                # Thực hiện ép mảng xuống Excel
                log_callback(f"[EXCEL] Tiến hành ép mảng RAM xuống vùng ô dữ liệu C7:Z{last_row}...")
                data_range.Value = grid_data
                log_callback("[EXCEL] Ép mảng thành công!")

                try:
                    ws.Protect(EXCEL_PASS)
                except:
                    pass

                log_callback(f"[EXCEL] Tiến hành Save file: {os.path.basename(excel_path)}")
                wb.Save()
                log_callback("[EXCEL] Lưu thành công!")

            except Exception as e:
                log_callback(f"[CRITICAL ERROR] Lỗi trong phiên làm việc của file Excel {os.path.basename(excel_path)}: {e}")

            finally:
                if wb:
                    try:
                        wb.Close(False)
                    except:
                        pass
                    del wb
                    gc.collect()

        if progress_callback:
            progress_callback(100)

    finally:
        if excel_app:
            excel_app.ScreenUpdating = True
            excel_app.DisplayAlerts = True
            excel_app.Quit()
            del excel_app

        pythoncom.CoUninitialize()
        gc.collect()

    return True, f"Hoàn thành tích hợp: {processed_count} mặt hàng."