import os
import sys
import re
from typing import Dict, List, Optional, Tuple

from app.logic.salary_config_parser import get_salary_config
from app.logic.metadata_manager import get_metadata_manager

EXCEL_PASS = os.getenv("EXCEL_SHEET_PASSWORD", "8863")
XL_CALC_MANUAL = -4135
EXCEL_RESTART_EVERY_N_FILES = 5


def get_config_key_for_source_file(source_file: str) -> Optional[str]:
    sf = source_file.lower()
    if "danh sách cbcnv bản dùng làm lương" in sf:
        return "danh_sach_cbcnv_luong"
    if "ck tháng" in sf or "chuyển khoản" in sf:
        return "chuyen_khoan_thang_truoc"
    if "danh sách cbcnv làm lương tháng" in sf:
        return "danh_sach_cbcnv_thang"
    if "chấm công" in sf:
        return "cham_cong"
    if "phụ cấp con nhỏ" in sf:
        return "phu_cap_con_nho"
    return None


def _column_letter_to_index(col_letter: str) -> int:
    col = 0
    for ch in col_letter.upper():
        col = col * 26 + (ord(ch) - ord("A") + 1)
    return col


def _column_index_to_letter(col_index: int) -> str:
    """Chuyển đổi chỉ số cột (1-based) thành ký tự (A, B, ..., Z, AA, AB, ...)"""
    result = ""
    while col_index > 0:
        col_index -= 1
        result = chr(65 + (col_index % 26)) + result
        col_index //= 26
    return result


def _resolve_source_sheet(
    source_sheet: str,
    resolved_source_sheet: str,
    sheet_mapping: dict,
) -> str:
    """Ánh xạ tên sheet từ config → sheet thực tế trong file user chọn."""
    if source_sheet in sheet_mapping:
        return sheet_mapping[source_sheet]
    if resolved_source_sheet in sheet_mapping:
        return sheet_mapping[resolved_source_sheet]
    # Template boc_vac cũ dùng "Sheet1" — file phụ cấp thực tế là "Con nhỏ"
    if resolved_source_sheet == "Sheet1" and "Con nhỏ" in sheet_mapping:
        return sheet_mapping["Con nhỏ"]
    return resolved_source_sheet


def _extract_team_name(file_path: str) -> Optional[str]:
    """
    Trích xuất tên tổ từ đường dẫn file (ví dụ: "sx/Cắt - 04.xlsx" → "Cắt").
    Hỗ trợ cả tên tiếng Việt và tiếng Anh.
    """
    basename = os.path.basename(file_path)
    # Remove extension and month suffix
    name_without_ext = os.path.splitext(basename)[0]  # "Cắt - 04"
    
    # Loại bỏ phần " - MM" ở cuối
    match = re.match(r'^(.+?)\s*-\s*\d{2}$', name_without_ext)
    if match:
        team_name = match.group(1).strip()
        return team_name
    
    return basename


def _match_team_config(team_name: str, team_configs: Dict) -> Optional[str]:
    """
    Tìm config phù hợp cho tổ từ team_configs.
    Hỗ trợ so khớp với pattern (ví dụ: "Cắt|Cut" để khớp với "Cắt" hoặc "Cut").
    Nếu không tìm thấy, trả về 'default' nếu có.
    """
    if not team_name or not team_configs:
        return "default" if "default" in team_configs else None
    
    team_name_lower = team_name.lower().strip()
    
    for pattern_key, config_value in team_configs.items():
        if pattern_key == "default":
            continue
        
        # Hỗ trợ pattern với "|" (ví dụ: "Cắt|Cut")
        patterns = [p.strip().lower() for p in pattern_key.split("|")]
        if team_name_lower in patterns:
            return pattern_key
    
    # Fallback to default
    return "default" if "default" in team_configs else None

class _ExcelDeploySession:
    """
    Ghi công thức bằng Excel COM.
    - Mở sẵn file nguồn (ReadOnly) → công thức dạng [tenfile.xlsx]Sheet (tránh treo vì path Unicode)
    - Không gọi Calculate khi lưu hàng loạt (tránh treo link ngoài)
    """

    def __init__(self):
        self.excel = None
        self._com_initialized = False
        self._source_book_names: Dict[str, str] = {}

    def start(self, log_callback, user_config: Optional[Dict] = None) -> bool:
        if sys.platform != "win32":
            log_callback("⚠️ Excel COM chỉ hỗ trợ Windows.")
            return False
        try:
            import pythoncom
            import win32com.client as win32

            try:
                pythoncom.CoInitializeEx(pythoncom.COINIT_APARTMENTTHREADED)
            except Exception:
                pythoncom.CoInitialize()
            self._com_initialized = True

            self.excel = win32.DispatchEx("Excel.Application")
            self.excel.Visible = False
            self.excel.DisplayAlerts = False
            self.excel.ScreenUpdating = False
            self.excel.AskToUpdateLinks = False
            self.excel.Interactive = False
            try:
                self.excel.EnableEvents = False
                self.excel.Calculation = XL_CALC_MANUAL
            except Exception:
                pass

            if user_config:
                self._preload_source_workbooks(user_config, log_callback)
                n = len(self._source_book_names)
                if n:
                    log_callback(f"📂 Đã mở sẵn {n} file nguồn (công thức ref ngắn, tránh treo)\n")
            return True
        except Exception as e:
            log_callback(f"❌ Không khởi tạo Excel: {e}")
            self.stop()
            return False

    def restart(self, log_callback, user_config: Optional[Dict] = None) -> bool:
        self.stop()
        return self.start(log_callback, user_config)

    @staticmethod
    def _pump_messages():
        try:
            import pythoncom
            pythoncom.PumpWaitingMessages()
        except Exception:
            pass

    def _preload_source_workbooks(self, user_config: Dict, log_callback):
        """Mở sẵn file nguồn ReadOnly — công thức không cần path đầy đủ (tránh treo)."""
        self._source_book_names.clear()
        seen = set()
        for cfg in user_config.values():
            path = (cfg or {}).get("file_path", "").strip()
            if not path:
                continue
            abs_path = os.path.normcase(os.path.abspath(path))
            if abs_path in seen or not os.path.isfile(abs_path):
                continue
            seen.add(abs_path)
            try:
                wb_src = self.excel.Workbooks.Open(
                    os.path.abspath(path), UpdateLinks=0, ReadOnly=True
                )
                self._source_book_names[abs_path] = wb_src.Name
            except Exception as e:
                log_callback(f"   ⚠️ Không mở trước file nguồn: {os.path.basename(path)} — {e}")

    def _build_vlookup_formula(
        self,
        current_row: int,
        source_path: str,
        actual_source_sheet: str,
        source_range: str,
        col_index: int,
    ) -> str:
        abs_path = os.path.normcase(os.path.abspath(source_path))
        book_name = self._source_book_names.get(abs_path)
        if book_name:
            # Dùng tên workbook ngắn cho file đã mở sẵn
            ref = f"'[{book_name}]{actual_source_sheet}'!{source_range}"
        else:
            # Dùng đường dẫn đầy đủ cho file chưa mở
            # Chuyển slash để tương thích với Excel
            abs_path_normalized = os.path.normpath(os.path.abspath(source_path))
            dir_path = os.path.dirname(abs_path_normalized)
            file_name = os.path.basename(source_path)
            # Format: 'C:\path\[filename]Sheet'!range
            ref = f"'{dir_path}\\[{file_name}]{actual_source_sheet}'!{source_range}"
        
        # Công thức: =IFERROR(VLOOKUP(C{row},ref,col,0),"")
        # Dùng "" (hai dấu ngoặc) cho chuỗi rỗng
        formula = f"=IFERROR(VLOOKUP(C{current_row},{ref},{col_index},0),\"\")"
        return formula

    def _write_formula_cell(self, ws, row: int, col: int, formula: str) -> bool:
        """
        Ghi công thức vào ô. Ưu tiên Formula2 chỉ cho formulas đơn giản.
        Với external references (VLOOKUP), dùng Formula thay vì Formula2.
        """
        if not formula.startswith("="):
            formula = f"={formula}"
        
        try:
            cell = ws.Cells(row, col)
        except Exception:
            return False
        
        # Mở khóa ô nếu cần
        try:
            cell.Locked = False
        except Exception:
            pass
        
        # Kiểm tra nếu formula có external reference (VLOOKUP với file khác)
        has_external_ref = "[" in formula and "]" in formula
        
        # Cách 1: Formula2 chỉ cho formulas đơn giản (không external refs)
        if not has_external_ref:
            try:
                cell.Formula2 = formula
                return True
            except Exception:
                pass
        
        # Cách 2: Thử chuyển đổi locale và dùng Formula (locale-aware)
        try:
            local = self.excel.ConvertFormula(formula, 1, 1, 1, None, True)
            if local:
                cell.Formula = local
                return True
        except Exception:
            pass
        
        # Cách 3: Trực tiếp gán Formula với cú pháp A1 tiếng Anh
        try:
            cell.Formula = formula
            return True
        except Exception:
            pass
        
        # Cách 4: Thử dùng Range thay vì Cells
        try:
            col_letter = _column_index_to_letter(col)
            cell_ref = f"{col_letter}{row}"
            range_obj = ws.Range(cell_ref)
            range_obj.Formula = formula
            return True
        except Exception:
            pass
        
        # Cách 5: Thử dùng Value (fallback cuối cùng)
        try:
            cell.Value = formula
            return True
        except Exception:
            pass
        
        return False

    def _open_target_workbook(self, file_path: str, log_callback):
        abs_path = os.path.normpath(os.path.abspath(file_path))
        if not os.path.isfile(abs_path):
            log_callback(f"   ❌ File không tồn tại: {abs_path}")
            return None
        try:
            return self.excel.Workbooks.Open(
                abs_path, UpdateLinks=0, ReadOnly=False, Password=EXCEL_PASS
            )
        except Exception:
            try:
                return self.excel.Workbooks.Open(abs_path, UpdateLinks=0, ReadOnly=False)
            except Exception as e:
                log_callback(
                    f"   ❌ Excel không mở được file:\n      {abs_path}\n"
                    f"      Lỗi: {e}\n      → Chạy lại Tab 1 để tạo file sạch từ template."
                )
                return None

    @staticmethod
    def _unprotect_workbook(wb):
        try:
            wb.Unprotect(Password=EXCEL_PASS)
        except Exception:
            try:
                wb.Unprotect()
            except Exception:
                pass

    @staticmethod
    def _unprotect_sheet(ws):
        for pwd in (EXCEL_PASS, ""):
            try:
                if pwd:
                    ws.Unprotect(Password=pwd)
                else:
                    ws.Unprotect()
            except Exception:
                pass

    @staticmethod
    def _find_sheet(wb, sheet_name: str):
        target = sheet_name.lower()
        for i in range(1, wb.Sheets.Count + 1):
            sh = wb.Sheets(i)
            if sh.Name.lower() == target:
                return sh
        return None

    def deploy_file(
        self,
        excel_file: str,
        template_key: str,
        year: int,
        month: int,
        config,
        log_callback,
        file_source_cache: Dict[str, Optional[str]],
        user_config: Optional[Dict] = None,
        team_name: Optional[str] = None,
    ) -> bool:
        if self.excel is None:
            return False

        yyyy, mm = int(year), int(month)
        mm_str = f"{mm:02d}"
        yy_short = str(yyyy)[-2:]
        if mm == 1:
            prev_month_int, prev_year_int = 12, yyyy - 1
        else:
            prev_month_int, prev_year_int = mm - 1, yyyy
        prev_month = f"{prev_month_int:02d}"
        prev_year = str(prev_year_int)
        prev_yy_short = prev_year[-2:]

        is_valid, errors = config.validate_template(template_key)
        if not is_valid:
            log_callback(f"   ❌ Template config lỗi: {errors[0]}")
            return False

        wb = self._open_target_workbook(excel_file, log_callback)
        if wb is None:
            return False

        file_modified = False
        template_config = config.get_template_config(template_key)
        file_lower = os.path.basename(excel_file).lower()
        bep_cong_vu_rate = "100%"
        if "bếp" in file_lower or "bep" in file_lower:
            bep_cong_vu_rate = "125%"
        elif "công vụ" in file_lower or "cong vu" in file_lower:
            bep_cong_vu_rate = "115%"

        try:
            self._unprotect_workbook(wb)

            for sheet_pattern, sheet_cfg in template_config.get("sheets", {}).items():
                resolved_sheet_name = sheet_pattern.replace("{mm}", mm_str)
                ws = self._find_sheet(wb, resolved_sheet_name)
                if ws is None:
                    log_callback(f"   ℹ️ Sheet '{resolved_sheet_name}' không tìm thấy → Bỏ qua")
                    continue

                self._unprotect_sheet(ws)
                start_row = sheet_cfg.get("start_row", 9)
                row_step = sheet_cfg.get("row_step", 2)

                num_employees = 0
                # Check cột B (tên) + C (mã NV), dừng nếu cả 2 trống (tối ưu tài nguyên)
                for row in range(start_row, start_row + 500, row_step):
                    val_b = ws.Cells(row, 2).Value
                    val_c = ws.Cells(row, 3).Value
                    str_b = str(val_b or "").strip()
                    str_c = str(val_c or "").strip()
                    
                    if str_b and str_c:  # Nếu cột B và C có dữ liệu
                        num_employees += 1
                    else:  # Cả 2 đều trống → dừng
                        break

                if num_employees == 0:
                    log_callback(f"   ℹ️ Sheet '{ws.Name}': Không có dữ liệu nhân viên")
                    continue

                log_callback(f"   📋 Sheet: {ws.Name} ({num_employees} nhân viên)")

                for col_letter, col_cfg in sheet_cfg.get("columns", {}).items():
                    parsed = config.parse_column_config(col_cfg)
                    col_num = _column_letter_to_index(col_letter)

                    if parsed["type"] == "team_formula":
                        team_configs = parsed.get("team_configs", {})
                        matched_key = _match_team_config(team_name, team_configs)
                        
                        if matched_key is None:
                            log_callback(f"      ⚠️ Cột {col_letter}: Không tìm thấy config cho tổ '{team_name}'")
                            continue
                        
                        config_value = team_configs[matched_key]
                        formula_template = config_value.get("formula", "")
                        desc = config_value.get("description", "")
                        
                        log_callback(f"      ✏️ Cột {col_letter}: {desc} (Tổ: {team_name or 'N/A'})")
                        fail_count = 0
                        for idx in range(num_employees):
                            current_row = start_row + (idx * row_step)
                            formula = formula_template.format(
                                row=current_row,
                                row_minus_1=current_row - 1,
                                row_plus_1=current_row + 1,
                                bep_cong_vu_rate=bep_cong_vu_rate,
                            )
                            if not self._write_formula_cell(ws, current_row, col_num, formula):
                                fail_count += 1
                            if idx % 10 == 9:
                                self._pump_messages()
                        if fail_count:
                            log_callback(
                                f"      ❌ Cột {col_letter}: {fail_count}/{num_employees} ô không ghi được"
                            )
                        else:
                            log_callback(f"         ↳ Đã ghi {num_employees} dòng cột {col_letter}")
                        file_modified = True

                    elif parsed["type"] == "formula":
                        formula_template = parsed["formula"]
                        log_callback(f"      ✏️ Cột {col_letter}: công thức nội bộ")
                        fail_count = 0
                        for idx in range(num_employees):
                            current_row = start_row + (idx * row_step)
                            formula = formula_template.format(
                                row=current_row,
                                row_minus_1=current_row - 1,
                                row_plus_1=current_row + 1,
                                bep_cong_vu_rate=bep_cong_vu_rate,
                            )
                            if not self._write_formula_cell(ws, current_row, col_num, formula):
                                fail_count += 1
                            if idx % 10 == 9:
                                self._pump_messages()
                        if fail_count:
                            log_callback(
                                f"      ❌ Cột {col_letter}: {fail_count}/{num_employees} ô không ghi được"
                            )
                        else:
                            log_callback(f"         ↳ Đã ghi {num_employees} dòng cột {col_letter}")
                        file_modified = True

                    elif parsed["type"] == "source":
                        source_file = parsed["file"]
                        source_sheet = parsed["sheet"]
                        source_range = parsed["range"]
                        col_index = parsed.get("column_index", 1)

                        resolved_file_name = source_file.format(
                            yyyy=yyyy, yy_short=yy_short, mm=mm_str, mm_int=mm,
                            prev_month=prev_month, prev_year=prev_year, prev_yy_short=prev_yy_short,
                        )
                        resolved_source_sheet = source_sheet.format(
                            yyyy=yyyy, yy_short=yy_short, mm=mm_str, mm_int=mm,
                            prev_month=prev_month, prev_year=prev_year, prev_yy_short=prev_yy_short,
                        )

                        source_path = None
                        actual_source_sheet = resolved_source_sheet
                        used_user_config = False
                        config_key = get_config_key_for_source_file(source_file)

                        if user_config and config_key and config_key in user_config:
                            uc = user_config[config_key]
                            user_file_path = uc.get("file_path", "")
                            if user_file_path and os.path.exists(user_file_path):
                                source_path = user_file_path
                                used_user_config = True
                                actual_source_sheet = _resolve_source_sheet(
                                    source_sheet,
                                    resolved_source_sheet,
                                    uc.get("sheet_mapping", {}),
                                )

                        if source_path is None:
                            cache_key = f"{source_file}:{yyyy}:{mm}"
                            if cache_key not in file_source_cache:
                                file_source_cache[cache_key] = config.resolve_file_path(
                                    source_file, yyyy, mm, lambda x: None
                                )
                            source_path = file_source_cache[cache_key]

                        if source_path is None:
                            log_callback(f"      ⚠️ Cột {col_letter}: '{resolved_file_name}' không tìm thấy")
                            for idx in range(num_employees):
                                ws.Cells(start_row + idx * row_step, col_num).Value = None
                            file_modified = True
                            continue

                        file_name = os.path.basename(source_path)
                        src_indicator = " (user)" if used_user_config else ""
                        short_ref = os.path.normcase(os.path.abspath(source_path)) in self._source_book_names
                        ref_note = " [ref ngắn]" if short_ref else " [ref path]"
                        log_callback(
                            f"      ✅ Cột {col_letter}: → {file_name} [{actual_source_sheet}]"
                            f"{src_indicator}{ref_note}"
                        )

                        fail_count = 0
                        for idx in range(num_employees):
                            current_row = start_row + (idx * row_step)
                            formula = self._build_vlookup_formula(
                                current_row,
                                source_path,
                                actual_source_sheet,
                                source_range,
                                col_index,
                            )
                            if not self._write_formula_cell(ws, current_row, col_num, formula):
                                fail_count += 1
                            if idx % 5 == 4:
                                self._pump_messages()

                        if fail_count:
                            log_callback(
                                f"      ❌ Cột {col_letter}: {fail_count}/{num_employees} ô không ghi được"
                            )
                        else:
                            log_callback(f"         ↳ Đã ghi {num_employees} dòng cột {col_letter}")
                        file_modified = True

            if file_modified:
                try:
                    self.excel.CalculateBeforeSave = False
                except Exception:
                    pass
                wb.Save()
                log_callback(
                    "   ✅ Đã lưu công thức\n"
                    "      (Mở file trong Excel — nhấn F9 hoặc Ctrl+Alt+F9 để tính kết quả)\n"
                )
            else:
                log_callback("   ℹ️ Không có thay đổi\n")
            return True

        except Exception as e:
            log_callback(f"   ❌ Lỗi xử lý: {e}")
            import traceback
            log_callback(traceback.format_exc())
            return False
        finally:
            try:
                if wb is not None:
                    wb.Close(SaveChanges=False)
            except Exception:
                pass
            self._pump_messages()

    def stop(self):
        try:
            if self.excel is not None:
                self.excel.Quit()
        except Exception:
            pass
        self.excel = None
        self._source_book_names.clear()
        if self._com_initialized:
            try:
                import pythoncom
                pythoncom.CoUninitialize()
            except Exception:
                pass
            self._com_initialized = False


def deploy_salary_from_metadata(
    year: int,
    month: int,
    log_callback,
    progress_callback=None,
    user_config: Dict = None,
) -> Tuple[bool, str]:
    try:
        config = get_salary_config()
    except Exception as e:
        return False, f"❌ Lỗi đọc config JSON: {str(e)}"

    metadata_mgr = get_metadata_manager()
    if not metadata_mgr.load_metadata(year, month):
        return False, f"❌ Không tìm thấy metadata.json tại: {metadata_mgr.metadata_path}"

    is_valid, errors = metadata_mgr.validate_metadata(log_callback)
    if not is_valid:
        return False, "❌ Metadata không hợp lệ:\n" + "\n".join(errors)

    log_callback("📊 TRIỂN KHAI CÔNG LƯƠNG - Từ Metadata")
    log_callback(f"📅 Năm/Tháng: {year}/{month:02d}")
    log_callback(f"📂 Metadata: {metadata_mgr.metadata_path}")
    if user_config:
        log_callback("⚙️ Sử dụng cấu hình nguồn do người dùng chọn")
    log_callback("")

    all_groups = metadata_mgr.get_all_groups_with_templates()
    if not all_groups:
        return False, "❌ Metadata không chứa thông tin file/template nào"

    log_callback(f"🔍 Tìm thấy {len(all_groups)} group nhân viên\n")

    session = _ExcelDeploySession()
    if not session.start(log_callback, user_config):
        return False, "❌ Không khởi tạo được Microsoft Excel."

    processed = 0
    warnings: List[str] = []
    file_source_cache: Dict[str, Optional[str]] = {}

    try:
        for idx, (relative_path, template_key, folder) in enumerate(all_groups):
            group_name = relative_path
            try:
                if template_key == "chua_xu_ly_template.xlsx":
                    log_callback(f"📄 [{idx + 1}/{len(all_groups)}] Bỏ qua: {relative_path}\n")
                    continue

                if (
                    idx > 0
                    and EXCEL_RESTART_EVERY_N_FILES > 0
                    and idx % EXCEL_RESTART_EVERY_N_FILES == 0
                ):
                    log_callback(f"♻️ [{idx + 1}/{len(all_groups)}] Khởi động lại Excel...\n")
                    if not session.restart(log_callback, user_config):
                        return False, "❌ Không thể khởi động lại Excel."

                mapping = metadata_mgr.get_split_files()
                group_name = mapping.get(relative_path, {}).get("group_name", relative_path)
                split_output = config.config.get(
                    "split_output_path", "D:\\Linh_Salary_Tool\\01_danh_sach_chia_to"
                )
                excel_file = os.path.join(
                    split_output, str(year), f"{int(month):02d}", relative_path
                )

                if not os.path.exists(excel_file):
                    msg = f"⚠️ [{idx + 1}/{len(all_groups)}] Không tìm thấy: {relative_path}"
                    log_callback(msg)
                    warnings.append(msg)
                    continue

                # Trích xuất tên tổ từ relative_path
                team_name = _extract_team_name(relative_path)

                log_callback(f"📄 [{idx + 1}/{len(all_groups)}] Xử lý: {group_name}")
                log_callback(f"   📋 Template: {template_key}")
                log_callback(f"   📁 File: {relative_path}\n")

                if session.deploy_file(
                    excel_file, template_key, year, month, config,
                    log_callback, file_source_cache, user_config, team_name=team_name,
                ):
                    processed += 1
                else:
                    warnings.append(f"⚠️ {relative_path}")

            except Exception as e:
                log_callback(f"❌ [{idx + 1}/{len(all_groups)}] {group_name}: {e}")
                warnings.append(str(e))
                import traceback
                log_callback(traceback.format_exc() + "\n")

            if progress_callback:
                progress_callback(int(((idx + 1) / len(all_groups)) * 100))
    finally:
        session.stop()

    if progress_callback:
        progress_callback(100)

    result_msg = f"\n✅ Hoàn tất! Đã triển khai {processed}/{len(all_groups)} file."
    if warnings:
        result_msg += f"\n⚠️ {len(warnings)} cảnh báo — xem log."
    return True, result_msg
