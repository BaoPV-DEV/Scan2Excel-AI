import gc
import json
import os
import re
import time
import pythoncom
import win32com.client as win32
from dotenv import load_dotenv


def normalize_str(s):
    if not s:
        return ""
    s = str(s).strip().lower()
    # Loại bỏ dấu tiếng Việt và chuẩn hóa ký tự
    s = re.sub(r'[àáạảãâầấậẩẫăằắặẳẵ]', 'a', s)
    s = re.sub(r'[èéẹẻẽêềếệểễ]', 'e', s)
    s = re.sub(r'[ìíịỉĩ]', 'i', s)
    s = re.sub(r'[òóọỏõôồốộổỗơờớợởỡ]', 'o', s)
    s = re.sub(r'[ùúụủũưừứựửữ]', 'u', s)
    s = re.sub(r'[ỳýỵỷỹ]', 'y', s)
    s = re.sub(r'[đ]', 'd', s)
    s = re.sub(r'\s+', ' ', s)
    return s

def teams_match(norm_json, norm_header):
    if norm_json == norm_header:
        return True
    # Ánh xạ đặc biệt (bí danh)
    if norm_header in ["co dong", "dieu dong"] and norm_json in ["co dong", "dieu dong"]:
        return True
    if norm_header in ["hoan thien", "h.thien"] and norm_json in ["hoan thien", "h.thien"]:
        return True
    return False

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
            try: ws_tong_hop.Unprotect()
            except: pass
            
            cells_to_format = ["R4", "S4", "T4"]
            for cell_ref in cells_to_format:
                cell = ws_tong_hop.Range(cell_ref)
                cell.HorizontalAlignment = -4108 # xlCenter
                cell.VerticalAlignment = -4108   # xlCenter
                cell.Font.Bold = True
                cell.WrapText = True
                cell.Interior.Color = 65535      # Yellow
                cell.Font.Color = 255            # Red
    except Exception:
        pass


def process_excel_integration(json_folder, excel_target_folder, log_callback, progress_callback=None, excel_pass="", single_dept=None):
    # Cấu hình mật khẩu
    load_dotenv()
    EXCEL_PASS = excel_pass or os.getenv("EXCEL_SHEET_PASSWORD", "8863")

    # 1. Thu thập và nhóm các file JSON theo file Excel đích
    json_paths = []
    for root, _, files in os.walk(json_folder):
        for f in files:
            if f.lower().endswith(".json"):
                json_paths.append(os.path.join(root, f))

    if not json_paths:
        return False, "Không tìm thấy file JSON nào."

    excel_groups = {}
    employee_dept_giay = {}
    for jp in json_paths:
        try:
            with open(jp, "r", encoding="utf-8") as f:
                data = json.load(f)
            info = data.get("thong_tin_chung", {})
            to_sx = str(info.get("to_san_xuat", "")).strip()
            if to_sx.isdigit(): to_sx = f"Tổ {to_sx}"
            to_sx_upper = to_sx.upper()
            
            # Tính toán Giây ngoài tổ cho tất cả file JSON (để fill sang các tổ khác)
            thuong = float(info.get("thuong_ma_hang_moi", 0)) / 100.0
            for cd in data.get("danh_sach_cong_doan", []):
                dinh_muc = float(cd.get("dinh_muc_t", 0))
                if dinh_muc <= 0: continue
                for th in cd.get("thuc_hien", []):
                    m_id = str(th.get("ma_nhan_vien", "")).strip().replace(".0", "")
                    sl = float(th.get("so_luong", 0))
                    if m_id and sl > 0:
                        if m_id not in employee_dept_giay:
                            employee_dept_giay[m_id] = {}
                        giay = (sl * dinh_muc / 100.0) * (1 + thuong)
                        employee_dept_giay[m_id][to_sx_upper] = employee_dept_giay[m_id].get(to_sx_upper, 0.0) + giay

            # Lọc theo single_dept nếu người dùng chọn chạy 1 tổ
            if single_dept and single_dept.upper() != to_sx_upper:
                continue

            match = re.search(r"(\d{2})/(\d{4})", str(info.get("thoi_gian", "")))
            if not match: continue
            
            mm = match.group(1)
            excel_name = f"{to_sx} - {mm}.xlsx"
            
            # Xác định thư mục 'sx' theo cách thông minh và chống lỗi
            sx_folder = excel_target_folder
            base_name = os.path.basename(excel_target_folder.rstrip("\\/")).lower()
            if base_name != "sx":
                potential_sx = os.path.join(excel_target_folder, "sx")
                if os.path.exists(potential_sx) or not os.path.exists(excel_target_folder):
                    sx_folder = potential_sx
                
            target_path = os.path.abspath(os.path.join(sx_folder, excel_name))
            
            if target_path not in excel_groups: excel_groups[target_path] = []
            excel_groups[target_path].append({"path": jp, "data": data, "to_sx_upper": to_sx_upper})
        except: continue

    # Lấy thông tin tháng và năm từ bất kỳ file JSON nào hợp lệ
    mm = None
    yyyy = None
    for jp in json_paths:
        try:
            with open(jp, "r", encoding="utf-8") as f:
                data = json.load(f)
            info = data.get("thong_tin_chung", {})
            match = re.search(r"(\d{2})/(\d{4})", str(info.get("thoi_gian", "")))
            if match:
                mm = match.group(1)
                yyyy = match.group(2)
                break
        except:
            continue

    # 2. Khởi tạo ứng dụng Excel
    pythoncom.CoInitialize()
    excel_app = None
    processed_count = 0
    total_files = len(json_paths)

    try:
        excel_app = win32.DispatchEx("Excel.Application")
        excel_app.Visible = False
        excel_app.DisplayAlerts = False  
        excel_app.ScreenUpdating = False 

        # Xác định tất cả các file Excel đích của tháng này (chỉ scan trong thư mục con 'sx' theo yêu cầu)
        import glob
        excel_files_in_folder = []
        if mm:
            # Xác định thư mục 'sx' theo cách thông minh và chống lỗi
            sx_folder = excel_target_folder
            base_name = os.path.basename(excel_target_folder.rstrip("\\/")).lower()
            if base_name != "sx":
                potential_sx = os.path.join(excel_target_folder, "sx")
                if os.path.exists(potential_sx) or not os.path.exists(excel_target_folder):
                    sx_folder = potential_sx
                
            pattern_sx = os.path.abspath(os.path.join(sx_folder, f"* - {mm}.xlsx"))
            excel_files_in_folder = [f for f in glob.glob(pattern_sx) if not os.path.basename(f).startswith("~$")]

        # Lọc nếu chạy cho 1 tổ
        if single_dept:
            single_dept_norm = normalize_str(single_dept)
            filtered_files = []
            for f in excel_files_in_folder:
                base = os.path.basename(f)
                to_sx_part = base.split(" - ")[0]
                if normalize_str(to_sx_part) == single_dept_norm:
                    filtered_files.append(f)
            excel_files_in_folder = filtered_files

        # Tập hợp tất cả các file Excel cần xử lý (gồm cả có JSON và không có JSON)
        # Chuẩn hóa đường dẫn để tránh sai khác chữ hoa/thường trên Windows
        normalized_excel_groups = {os.path.normpath(os.path.normcase(k)): v for k, v in excel_groups.items()}
        normalized_files_in_folder = {os.path.normpath(os.path.normcase(f)) for f in excel_files_in_folder}
        
        all_excel_paths = sorted(list(normalized_excel_groups.keys() | normalized_files_in_folder))

        # Khử trùng lặp: Nếu cùng 1 tên file xuất hiện ở cả thư mục con (sx/gt) và thư mục gốc,
        # chỉ giữ lại bản ghi trong thư mục con để cập nhật trực tiếp tại đó, tránh xử lý trùng.
        unique_paths = {}
        for path in all_excel_paths:
            filename = os.path.basename(path).lower()
            if filename not in unique_paths:
                unique_paths[filename] = path
            else:
                old_path = unique_paths[filename]
                old_is_sub = ("\\sx\\" in old_path.lower() or "/sx/" in old_path.lower() or 
                              "\\gt\\" in old_path.lower() or "/gt/" in old_path.lower())
                new_is_sub = ("\\sx\\" in path.lower() or "/sx/" in path.lower() or 
                              "\\gt\\" in path.lower() or "/gt/" in path.lower())
                if new_is_sub and not old_is_sub:
                    unique_paths[filename] = path
        
        all_excel_paths = sorted(list(unique_paths.values()))

        for idx, excel_path in enumerate(all_excel_paths):
            if not os.path.exists(excel_path):
                log_callback(f"❌ Không tìm thấy file: {os.path.basename(excel_path)}")
                continue

            wb = None
            # Trường hợp 1: File Excel có JSON sản lượng đi kèm
            if excel_path in normalized_excel_groups:
                items = normalized_excel_groups[excel_path]
                log_callback(f"\n📂 Bắt đầu xử lý file: {os.path.basename(excel_path)}")
                try:
                    wb = excel_app.Workbooks.Open(excel_path, UpdateLinks=0, ReadOnly=False, Password=EXCEL_PASS)
                except:
                    wb = excel_app.Workbooks.Open(excel_path)

                try:
                    # --- Mở khóa Workbook ---
                    try: wb.Unprotect(Password=EXCEL_PASS)
                    except:
                        try: wb.Unprotect()
                        except: pass

                    ws_source = wb.Sheets("Bang TH nop")
                    ws_template = wb.Sheets("xxx")

                    # =========================================================
                    # BƯỚC 1: LẤY DANH SÁCH TỔ CHUẨN (ƯU TIÊN SHEET "DANH SÁCH")
                    # =========================================================
                    log_callback("🔎 Đang xác định danh sách Tổ chuẩn...")
                    core_org_list = []
                    seen_ids_in_org = set()
                    
                    # Thử đọc từ sheet "Danh sách" trước
                    ws_ds = None
                    for s in wb.Sheets:
                        if s.Name == "Danh sách":
                            ws_ds = s
                            break
                    
                    if ws_ds:
                        log_callback("✅ Lấy danh sách từ sheet 'Danh sách'")
                        l_row_nv = ws_ds.Cells(ws_ds.Rows.Count, 4).End(-4162).Row
                        for r in range(2, l_row_nv + 1):
                            ma = str(ws_ds.Cells(r, 4).Value or "").strip().replace(".0", "")
                            ten = str(ws_ds.Cells(r, 5).Value or "").strip()
                            if ma and ten and ma not in seen_ids_in_org:
                                core_org_list.append({"ma": ma, "ten": ten})
                                seen_ids_in_org.add(ma)
                    else:
                        log_callback("ℹ️ Chưa có sheet 'Danh sách', lấy từ 'Bang TH nop'")
                        last_row_src = ws_source.Cells(ws_source.Rows.Count, 3).End(-4162).Row
                        for r in range(9, last_row_src + 1, 2):
                            ma = str(ws_source.Cells(r, 3).Value or "").strip().replace(".0", "")
                            ten = str(ws_source.Cells(r, 2).Value or "").strip()
                            if ten == "0": ten = ""
                            if ma and ten and len(ma) <= 20 and ma not in seen_ids_in_org:
                                core_org_list.append({"ma": ma, "ten": ten})
                                seen_ids_in_org.add(ma)

                    # Thu thập thông tin nhân viên từ JSON (CHỈ LẤY MÃ, BỎ TÊN NGOÀI TỔ)
                    all_involved_external_ids = set()
                    mh_involved_map = {} 
                    
                    for item in items:
                        data_mh = item["data"]
                        ma_h = str(data_mh.get("thong_tin_chung", {}).get("ma_hang", "UNKNOWN")).strip()
                        mh_involved_map[ma_h] = set()
                        
                        for cd in data_mh.get("danh_sach_cong_doan", []):
                            for th in cd.get("thuc_hien", []):
                                m_id = str(th.get("ma_nhan_vien", "")).strip().replace(".0", "")
                                if m_id and m_id.upper() != "N/A" and len(m_id) <= 20:
                                    if m_id not in seen_ids_in_org:
                                        all_involved_external_ids.add(m_id)
                                    mh_involved_map[ma_h].add(m_id)

                    # Master List duy nhất (Tổ có tên, Ngoài tổ để trống tên)
                    sorted_external = sorted(list(all_involved_external_ids))
                    final_full_list = core_org_list + [{"ma": m, "ten": ""} for m in sorted_external]
                    master_row_lookup = {nv["ma"]: 9 + idx*2 for idx, nv in enumerate(final_full_list)}

                    # =========================================================
                    # BƯỚC 2: XỬ LÝ CHI TIẾT TỪNG MÃ HÀNG
                    # =========================================================
                    mapping_all_mh = {} 
                    
                    for item in items:
                        data = item["data"]
                        info = data.get("thong_tin_chung", {})
                        ma_hang = str(info.get("ma_hang", "UNKNOWN")).strip()
                        log_callback(f"  > Đang xử lý mã hàng: {ma_hang}")

                        ext_ids_this_mh = [m for m in mh_involved_map.get(ma_hang, []) if m not in seen_ids_in_org]
                        local_nv_list = core_org_list + [{"ma": m, "ten": ""} for m in ext_ids_this_mh]
                        
                        safe_name = re.sub(r'[\\/*?:\[\]]', '_', ma_hang)[:31]
                        for s in wb.Sheets:
                            if s.Name.upper() == safe_name.upper():
                                try: s.Delete()
                                except: pass
                                break
                        
                        ws_new = wb.Sheets.Add(Before=wb.Sheets(1))
                        ws_template.Cells.Copy()
                        ws_new.Range("A1").PasteSpecial(-4104)
                        ws_new.Range("A1").PasteSpecial(8)
                        excel_app.CutCopyMode = False 
                        ws_new.Name = safe_name
                        ws_new.Tab.ColorIndex = 4

                        ws_new.Cells(1, 2).Value = ma_hang
                        ws_new.Cells(2, 2).Value = info.get("tong_san_luong_muc_tieu", 0)
                        ws_new.Cells(1, 3).Value = f"BẢNG LƯƠNG {info.get('thoi_gian', '')} Tổ {info.get('to_san_xuat', '')}".upper()

                        stt_col_map = {}
                        for cd in data.get("danh_sach_cong_doan", []):
                            stt = int(cd.get("stt", 0))
                            col = 5 + stt
                            ws_new.Cells(6, col).Value = cd.get("dinh_muc_t", 0)
                            stt_col_map[stt] = col

                        local_row_lookup = {}
                        r_curr = 9
                        for nv in local_nv_list:
                            ws_new.Cells(r_curr, 1).Value = (r_curr - 9) // 2 + 1
                            ws_new.Cells(r_curr, 2).Value = nv["ten"]
                            ws_new.Cells(r_curr, 3).Value = nv["ma"]
                            ws_new.Cells(r_curr, 4).Formula = f"=SUMPRODUCT(F{r_curr}:HE{r_curr},$F$6:$HE$6)/100"
                            local_row_lookup[nv["ma"]] = r_curr
                            r_curr += 2

                        for cd in data.get("danh_sach_cong_doan", []):
                            col = stt_col_map.get(int(cd.get("stt", 0)))
                            if not col: continue
                            for th in cd.get("thuc_hien", []):
                                m_id = str(th.get("ma_nhan_vien", "")).strip().replace(".0", "")
                                sl = th.get("so_luong", 0)
                                if m_id in local_row_lookup and sl > 0:
                                    tr = local_row_lookup[m_id]
                                    ws_new.Cells(tr, col).Value = (ws_new.Cells(tr, col).Value or 0) + sl
                                    ws_new.Cells(tr, col).NumberFormat = "#,##0"

                        mapping_all_mh[ma_hang] = local_row_lookup
                        processed_count += 1

                    # =========================================================
                    # BƯỚC 3: CẬP NHẬT SHEET TỔNG HỢP
                    # =========================================================
                    log_callback("📊 Đang cập nhật Sheet Tổng hợp...")
                    try:
                        ws_tong_hop = wb.Sheets("Tong hop cac ma")
                        
                        # Unprotect sheet trước khi sửa
                        try: ws_tong_hop.Unprotect(Password=EXCEL_PASS)
                        except: pass
                        
                        ws_tong_hop.Range("A9:Q1000").ClearContents()
                        # Không xóa trắng vùng Header (3-7) để giữ định dạng Merged Cells của R,S,T
                        
                        info_last = items[-1]["data"]["thong_tin_chung"]
                        ws_tong_hop.Cells(2, 1).Value = f"BẢNG TỔNG HỢP GIÂY {info_last.get('thoi_gian', '').upper()} - TỔ {str(info_last.get('to_san_xuat', '')).upper()}"
                        
                        for nv in final_full_list:
                            r_m = master_row_lookup[nv["ma"]]
                            ws_tong_hop.Cells(r_m, 1).Value = (r_m - 9) // 2 + 1
                            ws_tong_hop.Cells(r_m, 2).Value = nv["ten"]
                            ws_tong_hop.Cells(r_m, 3).Value = nv["ma"]

                        for idx, item in enumerate(items):
                            c_idx = 4 + idx
                            if c_idx > 17: break # Chỉ điền đến cột Q
                            info_mh = item["data"]["thong_tin_chung"]
                            mh_name = info_mh.get("ma_hang", "")
                            safe_n = re.sub(r'[\\/*?:\[\]]', '_', mh_name)[:31]
                            
                            ws_tong_hop.Cells(3, c_idx).Value = info_mh.get("thuong_ma_hang_moi", 0) / 100
                            ws_tong_hop.Cells(3, c_idx).NumberFormat = "0%"
                            ws_tong_hop.Cells(4, c_idx).Value = mh_name
                            ws_tong_hop.Cells(5, c_idx).Value = info_mh.get("tong_san_luong_muc_tieu", 0)
                            ws_tong_hop.Cells(7, c_idx).Value = f"M{idx+1}"
                            
                            local_map = mapping_all_mh.get(mh_name, {})
                            for nv in final_full_list:
                                r_m = master_row_lookup[nv["ma"]]
                                if nv["ma"] in local_map:
                                    ws_tong_hop.Cells(r_m, c_idx).Formula = f"='{safe_n}'!D{local_map[nv['ma']]}"
                                else:
                                    ws_tong_hop.Cells(r_m, c_idx).Value = 0

                        # Cột R, S, T giữ nguyên công thức có sẵn trong template
                        pass

                        
                    except Exception as ex:
                        log_callback(f"⚠️ Lỗi cập nhật tổng hợp: {ex}")

                    # =========================================================
                    # BƯỚC 4: CẬP NHẬT SHEET "BANG TH NOP"
                    # =========================================================
                    log_callback("📊 Đang cập nhật Sheet 'Bang TH nop'...")
                    try:
                        # Unprotect sheet trước khi sửa
                        try: ws_source.Unprotect(Password=EXCEL_PASS)
                        except: pass

                        # Đọc danh sách các cột từ F4 đến AC4 để ánh xạ
                        col_mapping = {}
                        for col in range(6, 30): # F (6) -> AC (29)
                            header_val = ws_source.Cells(4, col).Value
                            if header_val:
                                col_mapping[col] = normalize_str(header_val)

                        # Xác định tập hợp mã nhân viên thuộc core (chính thức) của tổ này
                        core_ids = {nv["ma"].upper(): nv for nv in core_org_list if nv.get("ma")}
                        external_ids = {m.upper() for m in sorted_external}

                        # Duyệt qua các dòng từ 9 đến 267, bước nhảy 2 dòng
                        for r in range(9, 268, 2):
                            ma_nv = str(ws_source.Cells(r, 3).Value or "").strip().replace(".0", "")
                            if not ma_nv:
                                # Nếu ô mã NV trống, xóa trắng các cột F:AC ở dòng này
                                for col in range(6, 30):
                                    ws_source.Cells(r, col).Value = None
                                continue
                            
                            ma_nv_upper = ma_nv.upper()
                            
                            # Nếu là nhân viên ngoài tổ (external), bỏ trống phạm vi F:AC
                            if ma_nv_upper in external_ids:
                                for col in range(6, 30):
                                    ws_source.Cells(r, col).Value = None
                                continue
                                
                            # Nếu là nhân viên chính thức trong tổ
                            if ma_nv_upper in core_ids or ma_nv_upper in [k.upper() for k in employee_dept_giay.keys()]:
                                # Tìm thông tin giây/sản lượng của nhân viên này
                                emp_giay_data = {}
                                for k, v in employee_dept_giay.items():
                                    if k.upper() == ma_nv_upper:
                                        emp_giay_data = v
                                        break
                                
                                for col in range(6, 30):
                                    norm_header = col_mapping.get(col)
                                    if not norm_header:
                                        ws_source.Cells(r, col).Value = None
                                        continue
                                    
                                    # Tính tổng giây từ các tổ trong employee_dept_giay khớp với cột này
                                    total_val = 0.0
                                    for json_to, val in emp_giay_data.items():
                                        norm_json = normalize_str(json_to)
                                        if teams_match(norm_json, norm_header):
                                            total_val += val
                                    
                                    if total_val > 0:
                                        ws_source.Cells(r, col).Value = total_val
                                        ws_source.Cells(r, col).NumberFormat = "#,##0"
                                    else:
                                        ws_source.Cells(r, col).Value = None
                            else:
                                # Trường hợp mã nhân viên không có trong danh sách nào, xóa trắng F:AC
                                for col in range(6, 30):
                                    ws_source.Cells(r, col).Value = None

                        
                    except Exception as ex:
                        log_callback(f"⚠️ Lỗi cập nhật Bang TH nop: {ex}")

                    # Đảm bảo định dạng cho các ô tiêu đề đã merge trước khi Save
                    ensure_merged_headers_format(wb)

                    wb.Save()
                    log_callback(f"✔ Hoàn tất và lưu file thành công.")
                except Exception as e:
                    log_callback(f"❌ Lỗi xử lý trong file {os.path.basename(excel_path)}: {e}")
                finally:
                    if wb is not None:
                        try: wb.Close()
                        except: pass

            # Trường hợp 2: File Excel KHÔNG CÓ JSON sản lượng đi kèm
            else:
                log_callback(f"\n📂 Bắt đầu xử lý file (Không có sản lượng): {os.path.basename(excel_path)}")
                try:
                    wb = excel_app.Workbooks.Open(excel_path, UpdateLinks=0, ReadOnly=False, Password=EXCEL_PASS)
                except:
                    wb = excel_app.Workbooks.Open(excel_path)

                try:
                    # --- Mở khóa Workbook ---
                    try: wb.Unprotect(Password=EXCEL_PASS)
                    except:
                        try: wb.Unprotect()
                        except: pass

                    ws_ds = None
                    ws_tong_hop = None
                    ws_source = None
                    for s in wb.Sheets:
                        if s.Name == "Danh sách":
                            ws_ds = s
                        elif s.Name == "Tong hop cac ma":
                            ws_tong_hop = s
                        elif s.Name == "Bang TH nop":
                            ws_source = s

                    # Chỉ xử lý các file dùng template có sheet Danh sách và Tong hop cac ma
                    if ws_ds and ws_tong_hop:
                        log_callback("🔎 Phát hiện sheet 'Danh sách' và 'Tong hop cac ma', tiến hành đồng bộ...")
                        
                        # 1. Đọc danh sách nhân viên từ sheet "Danh sách"
                        core_org_list = []
                        seen_ids_in_org = set()
                        l_row_nv = ws_ds.Cells(ws_ds.Rows.Count, 4).End(-4162).Row
                        for r in range(2, l_row_nv + 1):
                            ma = str(ws_ds.Cells(r, 4).Value or "").strip().replace(".0", "")
                            ten = str(ws_ds.Cells(r, 5).Value or "").strip()
                            if ma and ten and ma not in seen_ids_in_org:
                                core_org_list.append({"ma": ma, "ten": ten})
                                seen_ids_in_org.add(ma)

                        # 2. Cập nhật sheet "Tong hop cac ma"
                        try: ws_tong_hop.Unprotect(Password=EXCEL_PASS)
                        except: pass

                        ws_tong_hop.Range("A9:Q1000").ClearContents()
                        
                        to_sx_name = os.path.basename(excel_path).split(" - ")[0]
                        ws_tong_hop.Cells(2, 1).Value = f"BẢNG TỔNG HỢP GIÂY THÁNG {mm}/{yyyy} - {to_sx_name.upper()}"
                        
                        for idx, nv in enumerate(core_org_list):
                            r_m = 9 + idx * 2
                            ws_tong_hop.Cells(r_m, 1).Value = idx + 1
                            ws_tong_hop.Cells(r_m, 2).Value = nv["ten"]
                            ws_tong_hop.Cells(r_m, 3).Value = nv["ma"]


                        # 3. Cập nhật sheet "Bang TH nop"
                        if ws_source:
                            log_callback("📊 Đang cập nhật Sheet 'Bang TH nop'...")
                            try:
                                try: ws_source.Unprotect(Password=EXCEL_PASS)
                                except: pass

                                # Đọc danh sách các cột từ F4 đến AC4 để ánh xạ
                                col_mapping = {}
                                for col in range(6, 30):
                                    header_val = ws_source.Cells(4, col).Value
                                    if header_val:
                                        col_mapping[col] = normalize_str(header_val)

                                core_ids = {nv["ma"].upper(): nv for nv in core_org_list if nv.get("ma")}

                                # Duyệt qua các dòng từ 9 đến 267, bước nhảy 2 dòng
                                for r in range(9, 268, 2):
                                    ma_nv = str(ws_source.Cells(r, 3).Value or "").strip().replace(".0", "")
                                    if not ma_nv:
                                        for col in range(6, 30):
                                            ws_source.Cells(r, col).Value = None
                                        continue
                                    
                                    ma_nv_upper = ma_nv.upper()
                                    
                                    if ma_nv_upper in core_ids or ma_nv_upper in [k.upper() for k in employee_dept_giay.keys()]:
                                        emp_giay_data = {}
                                        for k, v in employee_dept_giay.items():
                                            if k.upper() == ma_nv_upper:
                                                emp_giay_data = v
                                                break
                                        
                                        for col in range(6, 30):
                                            norm_header = col_mapping.get(col)
                                            if not norm_header:
                                                ws_source.Cells(r, col).Value = None
                                                continue
                                            
                                            total_val = 0.0
                                            for json_to, val in emp_giay_data.items():
                                                norm_json = normalize_str(json_to)
                                                if teams_match(norm_json, norm_header):
                                                    total_val += val
                                            
                                            if total_val > 0:
                                                ws_source.Cells(r, col).Value = total_val
                                                ws_source.Cells(r, col).NumberFormat = "#,##0"
                                            else:
                                                ws_source.Cells(r, col).Value = None
                                    else:
                                        for col in range(6, 30):
                                            ws_source.Cells(r, col).Value = None


                            except Exception as ex:
                                log_callback(f"⚠️ Lỗi cập nhật Bang TH nop: {ex}")

                        # Đảm bảo định dạng cho các ô tiêu đề đã merge trước khi Save
                        ensure_merged_headers_format(wb)

                        wb.Save()
                        log_callback(f"✔ Hoàn tất và lưu file thành công.")
                    else:
                        log_callback("ℹ️ Bỏ qua vì không phải file sử dụng template Tổ may.")
                except Exception as e:
                    log_callback(f"❌ Lỗi xử lý trong file {os.path.basename(excel_path)}: {e}")
                finally:
                    if wb is not None:
                        try: wb.Close()
                        except: pass

            # Cập nhật tiến độ sau khi xử lý xong một file Excel
            if progress_callback:
                progress_callback(int(((idx + 1) / len(all_excel_paths)) * 100))

        if progress_callback:
            progress_callback(100)

    except Exception as e:
        return False, f"❌ Lỗi hệ thống: {str(e)}"
    finally:
        if excel_app:
            excel_app.ScreenUpdating = True
            excel_app.DisplayAlerts = True
            excel_app.Quit()
        pythoncom.CoUninitialize()
        gc.collect()
        
    return True, f"Xử lý thành công {processed_count} mã hàng."