import gc
import json
import os
import re
import time
import pythoncom
import win32com.client as win32

def process_excel_integration(json_folder, excel_target_folder, log_callback, progress_callback=None, excel_pass=""):
    # Cấu hình mật khẩu
    EXCEL_PASS = excel_pass or "8863"

    # 1. Thu thập và nhóm các file JSON theo file Excel đích
    json_paths = []
    for root, _, files in os.walk(json_folder):
        for f in files:
            if f.lower().endswith(".json"):
                json_paths.append(os.path.join(root, f))

    if not json_paths:
        return False, "Không tìm thấy file JSON nào."

    excel_groups = {}
    for jp in json_paths:
        try:
            with open(jp, "r", encoding="utf-8") as f:
                data = json.load(f)
            info = data.get("thong_tin_chung", {})
            to_sx = str(info.get("to_san_xuat", "")).strip()
            if to_sx.isdigit(): to_sx = f"Tổ {to_sx}"
            
            match = re.search(r"(\d{2})/(\d{4})", str(info.get("thoi_gian", "")))
            if not match: continue
            
            excel_name = f"{to_sx} - {match.group(1)}.xlsx"
            target_path = os.path.abspath(os.path.join(excel_target_folder, excel_name))
            
            if target_path not in excel_groups: excel_groups[target_path] = []
            excel_groups[target_path].append({"path": jp, "data": data})
        except: continue

    # 2. Khởi tạo ứng dụng Excel
    pythoncom.CoInitialize()
    excel_app = None
    processed_count = 0
    total_files = len(json_paths)

    try:
        # Dọn dẹp Excel treo trước khi chạy
        os.system("taskkill /f /im excel.exe >nul 2>&1")
        time.sleep(1)
        
        excel_app = win32.DispatchEx("Excel.Application")
        excel_app.Visible = False
        excel_app.DisplayAlerts = False  
        excel_app.ScreenUpdating = False 

        for excel_path, items in excel_groups.items():
            if not os.path.exists(excel_path):
                log_callback(f"❌ Không tìm thấy file: {os.path.basename(excel_path)}")
                continue

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
                # BƯỚC 1: LẤY DANH SÁCH TỔ CHUẨN (ƯU TIÊN SHEET "NHÂN VIÊN TỔ")
                # =========================================================
                log_callback("🔎 Đang xác định danh sách Tổ chuẩn...")
                core_org_list = []
                seen_ids_in_org = set()
                
                # Thử đọc từ sheet "Nhân viên tổ" trước
                ws_nv_goc = None
                for s in wb.Sheets:
                    if s.Name == "Nhân viên tổ":
                        ws_nv_goc = s
                        break
                
                if ws_nv_goc:
                    log_callback("✅ Lấy danh sách từ sheet 'Nhân viên tổ'")
                    l_row_nv = ws_nv_goc.Cells(ws_nv_goc.Rows.Count, 3).End(-4162).Row
                    for r in range(2, l_row_nv + 1):
                        ma = str(ws_nv_goc.Cells(r, 3).Value or "").strip().replace(".0", "")
                        ten = str(ws_nv_goc.Cells(r, 2).Value or "").strip()
                        if ma and ten and ma not in seen_ids_in_org:
                            core_org_list.append({"ma": ma, "ten": ten})
                            seen_ids_in_org.add(ma)
                else:
                    log_callback("ℹ️ Chưa có sheet 'Nhân viên tổ', lấy từ 'Bang TH nop'")
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
                    if progress_callback:
                        progress_callback(int((processed_count / total_files) * 100))

                # =========================================================
                # BƯỚC 3: CẬP NHẬT SHEET TỔNG HỢP
                # =========================================================
                log_callback("📊 Đang cập nhật Sheet Tổng hợp...")
                try:
                    ws_tong_hop = wb.Sheets("Tong hop cac ma")
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

                    for nv in final_full_list:
                        r_m = master_row_lookup[nv["ma"]]
                        ws_tong_hop.Cells(r_m, 18).Formula = f"=+SUMPRODUCT(D{r_m}:Q{r_m}*$D$3:$Q$3)"
                        ws_tong_hop.Cells(r_m, 19).Formula = f"=SUM(D{r_m}:R{r_m})"
                except Exception as ex:
                    log_callback(f"⚠️ Lỗi cập nhật tổng hợp: {ex}")

                # =========================================================
                # BƯỚC 4: ĐỒNG BỘ BANG TH NOP & LINK
                # =========================================================
                log_callback("🔗 Đang đồng bộ và link dữ liệu về Bang TH nop...")
                try:
                    # 1. CHỈ xóa cột A-Q, tuyệt đối không chạm vào R,S,T (cột 18 trở đi)
                    ws_source.Range("A9:Q1000").ClearContents()
                    curr_r_src = 9
                    for idx, nv in enumerate(final_full_list):
                        if idx > 0:
                            try:
                                ws_source.Rows("9:10").Copy()
                                ws_source.Rows(f"{curr_r_src}:{curr_r_src+1}").Insert()
                                excel_app.CutCopyMode = False
                            except: pass

                        # 2. Điền thông tin STT, Tên, Mã
                        ws_source.Cells(curr_r_src, 1).Value = idx + 1
                        ws_source.Cells(curr_r_src, 2).Value = nv["ten"]
                        ws_source.Cells(curr_r_src, 3).Value = nv["ma"]
                        
                        # 3. Xóa dữ liệu cũ cột D-Q (cột 4-17), giữ lại công thức R,S,T
                        ws_source.Range(ws_source.Cells(curr_r_src, 4), ws_source.Cells(curr_r_src, 17)).ClearContents()
                        
                        # 4. Link dữ liệu từ Tong hop cac ma sang
                        r_m = master_row_lookup[nv["ma"]]
                        ws_source.Cells(curr_r_src, 5).Formula = f"='Tong hop cac ma'!S{r_m}"
                        curr_r_src += 2
                except Exception as ex:
                    log_callback(f"⚠️ Lỗi đồng bộ Bang TH nop: {ex}")

                # =========================================================
                # BƯỚC 5: TẠO SHEET "NHÂN VIÊN TỔ" (DANH SÁCH GỐC)
                # =========================================================
                log_callback("📋 Đang cập nhật danh sách Nhân viên tổ...")
                try:
                    sheet_name_nv = "Nhân viên tổ"
                    
                    # 1. Xóa sheet cũ nếu tồn tại
                    for s in wb.Sheets:
                        if s.Name == sheet_name_nv:
                            excel_app.DisplayAlerts = False
                            s.Delete()
                            excel_app.DisplayAlerts = True
                            time.sleep(0.5) # Đợi một chút để Excel cập nhật cấu trúc
                            break
                    
                    # 2. Thêm sheet mới vào vị trí cuối cùng
                    try:
                        # Thêm vào sau sheet cuối cùng hiện tại
                        ws_nv = wb.Worksheets.Add(After=wb.Sheets(wb.Sheets.Count))
                        ws_nv.Name = sheet_name_nv
                    except:
                        # Nếu lỗi thì thêm đại rồi Move sau
                        ws_nv = wb.Worksheets.Add()
                        ws_nv.Name = sheet_name_nv
                        try:
                            ws_nv.Move(After=wb.Sheets(wb.Sheets.Count))
                        except:
                            pass
                    
                    # 3. Header: TT, Họ và tên, Mã số
                    headers = ["TT", "Họ và tên", "Mã số"]
                    for i, h in enumerate(headers):
                        cell = ws_nv.Cells(1, i + 1)
                        cell.Value = h
                        cell.Interior.Color = 65535 # Yellow
                        cell.Font.Bold = True
                        cell.Borders.LineStyle = 1
                    
                    # Fill dữ liệu
                    for idx, nv in enumerate(core_org_list):
                        row = idx + 2
                        ws_nv.Cells(row, 1).Value = idx + 1
                        ws_nv.Cells(row, 2).Value = nv["ten"]
                        ws_nv.Cells(row, 3).Value = nv["ma"]
                        # Border cho data
                        for col in range(1, 4):
                            ws_nv.Cells(row, col).Borders.LineStyle = 1
                    
                    ws_nv.Columns("A:C").AutoFit()
                except Exception as ex:
                    log_callback(f"⚠️ Lỗi tạo sheet Nhân viên tổ: {ex}")

                wb.Save()
                log_callback(f"✔ Hoàn tất và lưu file thành công.")
            except Exception as e:
                log_callback(f"❌ Lỗi xử lý trong file {os.path.basename(excel_path)}: {e}")
            finally:
                wb.Close()

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