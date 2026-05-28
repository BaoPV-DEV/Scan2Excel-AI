metadata_manager.py:
import os
import json
from typing import Dict, List, Optional, Tuple
from datetime import datetime


class MetadataManager:
    """
    Quản lý metadata.json từ Tab 2 - tách nhân viên.
    Đọc metadata để xác định template và file đầu vào cho Tab 5.
    
    Metadata structure từ Tab 2:
    {
        "year_month": "2026-04",
        "split_date": "2026-04-15T10:30:00",
        "input_files": {
            "gt/Bảo vệ": {
                "output_file": "gt/Bảo vệ - 04.xlsx",
                "template": "gt/bao_ve_template.xlsx",
                "employee_count": 5
            }
        }
    }
    """
    
    def __init__(self, metadata_path: str = None):
        """
        Initialize metadata manager.
        
        Args:
            metadata_path: Đường dẫn đến metadata.json (mặc định: config value)
        """
        self.metadata_path = metadata_path
        self.metadata = {}
    
    def load_metadata(self, year: int, month: int) -> bool:
        """
        Tải metadata.json từ đường dẫn: 
        {split_output_path}\{YYYY}\{MM}\data\metadata.json
        
        Args:
            year: Năm (VD: 2026)
            month: Tháng (VD: 4)
            
        Returns:
            True nếu thành công, False nếu không
        """
        from logic.salary_config_parser import get_salary_config
        
        if self.metadata_path is None:
            # Xây dựng đường dẫn từ config
            config = get_salary_config()
            split_output = config.config.get("split_output_path", "D:\\Linh_Salary_Tool\\01_danh_sach_chia_to")
            mm = f"{int(month):02d}"
            yyyy = str(year)
            
            self.metadata_path = os.path.join(split_output, yyyy, mm, "data", "metadata.json")
        
        if not os.path.exists(self.metadata_path):
            return False
        
        try:
            with open(self.metadata_path, 'r', encoding='utf-8') as f:
                self.metadata = json.load(f)
            return True
        except Exception as e:
            print(f"❌ Lỗi đọc metadata.json: {e}")
            return False
    
    def get_year_month(self) -> Optional[str]:
        """
        Lấy năm-tháng từ metadata (VD: "2026-04")
        
        Returns:
            String "YYYY-MM" hoặc None nếu metadata chưa tải
        """
        return self.metadata.get("year_month")
    
    def get_split_files(self) -> Dict[str, Dict]:
        """
        Lấy danh sách file tách từ metadata.
        
        Returns:
            Dict với key=relative_path, value={template, group_name, employee_count, folder}
        """
        # Metadata từ Tab 1 dùng key "mapping" (không phải "input_files")
        mapping = self.metadata.get("mapping", {})
        return mapping
    
    def get_output_file_path(self, year: int, month: int, relative_path: str) -> Optional[str]:
        """
        Xây dựng đường dẫn đầy đủ tới file output từ Tab 1.
        
        Args:
            year: Năm
            month: Tháng
            relative_path: Đường dẫn tương đối (VD: "sx/Tổ 1 - 04.xlsx")
            
        Returns:
            Đường dẫn đầy đủ hoặc None nếu không tìm thấy
        """
        from logic.salary_config_parser import get_salary_config
        
        config = get_salary_config()
        split_output = config.config.get("split_output_path", "D:\\Linh_Salary_Tool\\01_danh_sach_chia_to")
        mm = f"{int(month):02d}"
        yyyy = str(year)
        
        # Lấy output_dir từ metadata
        output_dir = self.metadata.get("output_dir")
        if not output_dir:
            # Fallback: tạo từ year/month
            output_dir = os.path.join(split_output, yyyy, mm)
        
        full_path = os.path.join(output_dir, relative_path)
        return full_path if os.path.exists(full_path) else None
    
    def get_template_for_group(self, relative_path: str) -> Optional[str]:
        """
        Lấy template tương ứng với file từ metadata.
        
        Args:
            relative_path: Đường dẫn tương đối (VD: "sx/Tổ 1 - 04.xlsx")
            
        Returns:
            Tên template (VD: "sx/to_may_template.xlsx") hoặc None
        """
        mapping = self.get_split_files()
        if relative_path not in mapping:
            return None
        
        return mapping[relative_path].get("template")
    
    def get_all_groups_with_templates(self) -> List[Tuple[str, str, str]]:
        """
        Lấy danh sách tất cả file cùng templates.
        
        Returns:
            List của tuple (relative_path, template, folder)
            VD: [("sx/Tổ 1 - 04.xlsx", "sx/to_may_template.xlsx", "sx"), ...]
        """
        result = []
        mapping = self.get_split_files()
        
        for relative_path, file_info in mapping.items():
            template = file_info.get("template")
            folder = file_info.get("folder", "root")
            if template:
                result.append((relative_path, template, folder))
        
        return result
    
    def validate_metadata(self, log_callback=None) -> Tuple[bool, List[str]]:
        """
        Kiểm tra tính hợp lệ của metadata.
        
        Returns:
            (is_valid, list_of_errors)
        """
        errors = []
        
        if not self.metadata:
            errors.append("❌ Metadata trống - chưa tải file hoặc file không tồn tại")
            return False, errors
        
        if "year_month" not in self.metadata:
            errors.append("❌ Thiếu 'year_month' trong metadata")
        
        # Kiểm tra "mapping" (không phải "input_files")
        mapping = self.metadata.get("mapping", {})
        if not mapping:
            errors.append("❌ Không có 'mapping' trong metadata")
        else:
            for relative_path, file_info in mapping.items():
                if "template" not in file_info:
                    errors.append(f"⚠️ File '{relative_path}': thiếu 'template'")
        
        if log_callback and errors:
            for error in errors:
                log_callback(error)
        
        return len(errors) == 0, errors


# Singleton instance
_global_metadata = None

def get_metadata_manager(metadata_path: str = None) -> MetadataManager:
    """Get or create global metadata manager instance"""
    global _global_metadata
    if _global_metadata is None:
        _global_metadata = MetadataManager(metadata_path)
    return _global_metadata


salary_config_parser.py:
import os
import json
import re
from typing import Dict, List, Optional, Tuple

class SalarySourceConfig:
    """
    Parser for salary_sources.json configuration file.
    Manages template-to-source mapping and handles dynamic path resolution.
    """
    
    def __init__(self, config_path: str = None):
        """Initialize config parser from JSON file"""
        if config_path is None:
            # Default path: ../Template/salary_sources.json (one level up from logic/)
            current_dir = os.path.abspath(os.path.dirname(__file__))
            config_path = os.path.join(current_dir, "..", "Template", "salary_sources.json")
        
        self.config_path = config_path
        self.config = self._load_config()
        self.base_path = self.config.get("base_path", "D:\\Linh_Salary_Tool")
        self.templates = self.config.get("templates", {})
    
    def _load_config(self) -> dict:
        """Load JSON configuration file"""
        try:
            with open(self.config_path, 'r', encoding='utf-8') as f:
                return json.load(f)
        except FileNotFoundError:
            raise FileNotFoundError(f"Config file not found: {self.config_path}")
        except json.JSONDecodeError as e:
            raise ValueError(f"Invalid JSON format: {e}")
    
    def get_template_config(self, template_key: str) -> Optional[Dict]:
        """Get configuration for a specific template"""
        return self.templates.get(template_key)
    
    def resolve_file_path(self, file_pattern: str, year: int, month: int, 
                          log_callback=None) -> Optional[str]:
        """
        Resolve dynamic file path with year/month placeholders.
        Returns full path if file exists, None otherwise.
        """
        # Replace placeholders
        yyyy = str(year)
        mm = f"{int(month):02d}"
        yy_short = yyyy[-2:]
        mm_int = int(month)
        
        if mm_int == 1:
            prev_month_int = 12
            prev_year_int = year - 1
        else:
            prev_month_int = mm_int - 1
            prev_year_int = year
        
        prev_month = f"{prev_month_int:02d}"
        prev_year = str(prev_year_int)
        
        resolved_file = file_pattern.format(
            yyyy=yyyy,
            yy_short=yy_short,
            mm=mm,
            mm_int=mm_int,
            prev_month=prev_month,
            prev_year=prev_year,
            prev_yy_short=prev_year[-2:]
        )
        
        # Build full path: base_path\YYYY\MM\resolved_file
        full_path = os.path.join(
            self.base_path,
            str(year),
            mm,
            resolved_file
        )
        
        # Check if file exists
        if os.path.exists(full_path):
            return full_path
        else:
            if log_callback:
                log_callback(f"⚠️ File not found: {full_path}")
            return None
    
    def get_sheet_config(self, template_key: str, sheet_name: str) -> Optional[Dict]:
        """Get configuration for a specific sheet in a template"""
        template = self.get_template_config(template_key)
        if not template:
            return None
        
        sheets = template.get("sheets", {})
        
        # Handle dynamic sheet names like "Lương {mm}"
        for sheet_pattern, config in sheets.items():
            if sheet_pattern == sheet_name or sheet_name.lower().startswith(sheet_pattern.lower()):
                return config
        
        return None
    
    def get_column_sources(self, template_key: str, sheet_name: str) -> Dict:
        """Get all source file mappings for columns in a sheet"""
        sheet_config = self.get_sheet_config(template_key, sheet_name)
        if not sheet_config:
            return {}
        
        return sheet_config.get("columns", {})
    
    def parse_column_config(self, col_config: Dict) -> Dict:
        """
        Parse individual column configuration.
        Returns dict with 'type' (formula or source), and corresponding data.
        """
        if "formula" in col_config:
            return {
                "type": "formula",
                "formula": col_config["formula"],
                "description": col_config.get("description", "")
            }
        elif "source_file" in col_config:
            return {
                "type": "source",
                "file": col_config["source_file"],
                "sheet": col_config["source_sheet"],
                "range": col_config.get("source_range", ""),
                "column_index": col_config.get("column_index"),
                "description": col_config.get("description", "")
            }
        else:
            return {"type": "unknown"}
    
    def validate_template(self, template_key: str) -> Tuple[bool, List[str]]:
        """Validate template configuration. Returns (is_valid, list_of_errors)"""
        errors = []
        
        template = self.get_template_config(template_key)
        if not template:
            return False, [f"Template not found: {template_key}"]
        
        sheets = template.get("sheets", {})
        for sheet_name, sheet_config in sheets.items():
            columns = sheet_config.get("columns", {})
            for col_letter, col_config in columns.items():
                parsed = self.parse_column_config(col_config)
                if parsed["type"] == "unknown":
                    errors.append(f"Sheet '{sheet_name}', Column {col_letter}: Invalid config format")
        
        return len(errors) == 0, errors


# Utility function to get global config instance
_global_config = None

def get_salary_config(config_path: str = None) -> SalarySourceConfig:
    """Get or create global salary config instance"""
    global _global_config
    if _global_config is None:
        _global_config = SalarySourceConfig(config_path)
    return _global_config

salary_deployer.py:
import os
import openpyxl
from typing import Dict, Optional, Tuple
from logic.salary_config_parser import get_salary_config
from logic.metadata_manager import get_metadata_manager


def deploy_salary_from_metadata(year: int, month: int, log_callback, progress_callback=None) -> Tuple[bool, str]:
    """
    ✨ NEW: Triển khai công lương từ metadata - Tuần tự + Tối ưu hóa
    
    Workflow:
    1. Đọc metadata.json từ D:\Linh_Salary_Tool\01_danh_sach_chia_to\YYYY\MM\data\metadata.json
    2. Từ metadata xác định: file output, template tương ứng
    3. Tuần tự xử lý từng file:
       - Load file Excel
       - Áp dụng công thức từ config
       - Linking dữ liệu từ source files
       - Lưu file
    4. Tối ưu hóa: Cache file sources, batch operations
    
    Args:
        year: Năm (VD: 2026)
        month: Tháng (VD: 4)
        log_callback: Callback để ghi log
        progress_callback: Callback để cập nhật tiến độ (0-100)
        
    Returns:
        (success: bool, message: str)
    """
    try:
        config = get_salary_config()
    except Exception as e:
        return False, f"❌ Lỗi đọc config JSON: {str(e)}"
    
    # Bước 1: Tải metadata
    metadata_mgr = get_metadata_manager()
    if not metadata_mgr.load_metadata(year, month):
        return False, f"❌ Không tìm thấy metadata.json tại: {metadata_mgr.metadata_path}"
    
    # Kiểm tra tính hợp lệ metadata
    is_valid, errors = metadata_mgr.validate_metadata(log_callback)
    if not is_valid:
        return False, "❌ Metadata không hợp lệ:\n" + "\n".join(errors)
    
    log_callback(f"📊 TRIỂN KHAI CÔNG LƯƠNG - Từ Metadata")
    log_callback(f"📅 Năm/Tháng: {year}/{month:02d}")
    log_callback(f"📂 Metadata: {metadata_mgr.metadata_path}\n")
    
    # Bước 2: Lấy danh sách file và template từ metadata
    all_groups = metadata_mgr.get_all_groups_with_templates()
    if not all_groups:
        return False, "❌ Metadata không chứa thông tin file/template nào"
    
    log_callback(f"🔍 Tìm thấy {len(all_groups)} group nhân viên\n")
    
    # Bước 3: Tuần tự xử lý từng file
    processed = 0
    warnings = []
    
    # Cache để tối ưu hiệu năng: lưu file sources đã resolve
    file_source_cache: Dict[str, Optional[str]] = {}
    
    for idx, (group_name, template_key, output_filename) in enumerate(all_groups):
        try:
            # Xây dựng đường dẫn đầy đủ file output
            split_output = config.config.get("split_output_path", "D:\\Linh_Salary_Tool\\01_danh_sach_chia_to")
            mm = f"{int(month):02d}"
            yyyy = str(year)
            excel_file = os.path.join(split_output, yyyy, mm, output_filename)
            
            if not os.path.exists(excel_file):
                warning = f"⚠️ [{idx + 1}/{len(all_groups)}] File không tìm thấy: {output_filename}"
                log_callback(warning)
                warnings.append(warning)
                continue
            
            log_callback(f"📄 [{idx + 1}/{len(all_groups)}] Xử lý: {group_name}")
            log_callback(f"   📋 Template: {template_key}")
            log_callback(f"   📁 File: {output_filename}\n")
            
            # Xử lý file này
            success = _process_excel_file(
                excel_file,
                template_key,
                year,
                month,
                config,
                log_callback,
                file_source_cache
            )
            
            if success:
                processed += 1
            else:
                warnings.append(f"⚠️ File {output_filename}: Xử lý có lỗi")
            
        except Exception as e:
            warning = f"❌ [{idx + 1}/{len(all_groups)}] Lỗi xử lý group '{group_name}': {str(e)}"
            log_callback(warning)
            warnings.append(warning)
            import traceback
            log_callback(f"   {traceback.format_exc()}\n")
        
        # Cập nhật tiến độ
        if progress_callback:
            progress_callback(int(((idx + 1) / len(all_groups)) * 100))
    
    if progress_callback:
        progress_callback(100)
    
    # Kết quả
    result_msg = f"\n✅ Hoàn tất! Đã triển khai {processed}/{len(all_groups)} file."
    if warnings:
        result_msg += f"\n⚠️ Có {len(warnings)} cảnh báo - Kiểm tra log"
    
    return True, result_msg


def _process_excel_file(
    excel_file: str,
    template_key: str,
    year: int,
    month: int,
    config,
    log_callback,
    file_source_cache: Dict[str, Optional[str]]
) -> bool:
    """
    Xử lý một file Excel duy nhất.
    Tối ưu: Sử dụng cache cho file sources.
    
    Returns:
        True nếu thành công, False nếu lỗi
    """
    yyyy = int(year)
    mm = int(month)
    yyyy_str = str(yyyy)
    mm_str = f"{mm:02d}"
    mm_int = mm
    
    # Tính tháng trước
    if mm_int == 1:
        prev_month_int = 12
        prev_year_int = yyyy - 1
    else:
        prev_month_int = mm_int - 1
        prev_year_int = yyyy
    
    prev_month = f"{prev_month_int:02d}"
    prev_year = str(prev_year_int)
    yy_short = yyyy_str[-2:]
    prev_yy_short = prev_year[-2:]
    
    # Validate template
    is_valid, errors = config.validate_template(template_key)
    if not is_valid:
        log_callback(f"   ❌ Template config lỗi: {errors[0]}")
        return False
    
    try:
        wb = openpyxl.load_workbook(excel_file)
    except Exception as e:
        log_callback(f"   ❌ Lỗi mở file Excel: {str(e)}")
        return False
    
    file_modified = False
    template_config = config.get_template_config(template_key)
    
    try:
        # Xác định tỷ lệ cho Bếp/Công vụ (dựa trên tên file)
        file_lower = os.path.basename(excel_file).lower()
        bep_cong_vu_rate = "100%"
        if "bếp" in file_lower or "bep" in file_lower:
            bep_cong_vu_rate = "125%"
        elif "công vụ" in file_lower or "cong vu" in file_lower:
            bep_cong_vu_rate = "115%"
        
        # Lặp qua tất cả sheets trong template config
        for sheet_pattern, sheet_cfg in template_config.get("sheets", {}).items():
            # Resolve dynamic sheet name
            resolved_sheet_name = sheet_pattern.replace("{mm}", mm_str)
            
            # Tìm sheet trong workbook (case-insensitive)
            ws = None
            actual_sheet_name = None
            for name in wb.sheetnames:
                if name.lower() == resolved_sheet_name.lower():
                    ws = wb[name]
                    actual_sheet_name = name
                    break
            
            if not ws:
                log_callback(f"   ℹ️ Sheet '{resolved_sheet_name}' không tìm thấy → Bỏ qua")
                continue
            
            # Mở khóa sheet nếu có bảo vệ
            try:
                ws.protection.sheet = False
            except Exception:
                pass
            
            start_row = sheet_cfg.get("start_row", 9)
            row_step = sheet_cfg.get("row_step", 2)
            
            # Tính số lượng nhân viên từ dữ liệu trong cột C (mã nhân viên)
            num_employees = 0
            for row in range(start_row, start_row + 500, row_step):
                cell_value = ws.cell(row=row, column=3).value
                if cell_value and str(cell_value).strip():
                    num_employees += 1
                else:
                    break
            
            if num_employees == 0:
                log_callback(f"   ℹ️ Sheet '{actual_sheet_name}': Không có dữ liệu nhân viên")
                continue
            
            log_callback(f"   📋 Sheet: {actual_sheet_name} ({num_employees} nhân viên)")
            
            # Lặp qua từng cột có config
            columns_config = sheet_cfg.get("columns", {})
            for col_letter, col_cfg in columns_config.items():
                parsed = config.parse_column_config(col_cfg)
                col_num = ord(col_letter.upper()) - ord('A') + 1
                
                if parsed["type"] == "formula":
                    # Công thức tĩnh - áp dụng trực tiếp
                    formula_template = parsed["formula"]
                    
                    for idx in range(num_employees):
                        current_row = start_row + (idx * row_step)
                        formula = formula_template.format(
                            row=current_row,
                            row_minus_1=current_row - 1,
                            row_plus_1=current_row + 1,
                            bep_cong_vu_rate=bep_cong_vu_rate
                        )
                        try:
                            ws.cell(row=current_row, column=col_num).value = formula
                        except Exception as e:
                            log_callback(f"      ❌ Cột {col_letter}, dòng {current_row}: {str(e)}")
                    file_modified = True
                
                elif parsed["type"] == "source":
                    # Linking từ file ngoài - Sử dụng CACHE
                    source_file = parsed["file"]
                    source_sheet = parsed["sheet"]
                    source_range = parsed["range"]
                    col_index = parsed.get("column_index", 1)
                    
                    # Resolve file source (với cache)
                    cache_key = f"{source_file}:{yyyy}:{mm}"
                    if cache_key not in file_source_cache:
                        file_source_cache[cache_key] = config.resolve_file_path(
                            source_file, yyyy, mm, lambda x: None  # Silent log
                        )
                    
                    source_path = file_source_cache[cache_key]
                    
                    if source_path is None:
                        # File không tồn tại → để trống + warning
                        log_callback(f"      ⚠️ Cột {col_letter}: File '{source_file}' không tìm thấy")
                        
                        for idx in range(num_employees):
                            current_row = start_row + (idx * row_step)
                            ws.cell(row=current_row, column=col_num).value = None
                        file_modified = True
                    else:
                        # File tồn tại → tạo formula linking
                        for idx in range(num_employees):
                            current_row = start_row + (idx * row_step)
                            # Excel external link formula
                            formula = f"=IFERROR(VLOOKUP(C{current_row},'{source_path}[{os.path.basename(source_path)}]{source_sheet}'!{source_range},{col_index},0),\"\")"
                            try:
                                ws.cell(row=current_row, column=col_num).value = formula
                            except Exception as e:
                                log_callback(f"      ❌ Cột {col_letter}, dòng {current_row}: {str(e)}")
                        file_modified = True
        
        # Lưu file nếu có thay đổi
        if file_modified:
            wb.save(excel_file)
            log_callback(f"   ✅ Đã lưu file\n")
        else:
            log_callback(f"   ℹ️ Không có thay đổi\n")
        
        return True
    
    except Exception as e:
        log_callback(f"   ❌ Lỗi xử lý: {str(e)}")
        import traceback
        log_callback(f"   {traceback.format_exc()}")
        return False
    
    finally:
        wb.close()

Template\salary_sources.json:
{
    "config_version": "2.1",
    "description": "Salary deployment configuration - Tab 5 reads metadata from Tab 2 to auto-detect templates and deploy formulas",
    "metadata": {
        "source_path": "D:\\Linh_Salary_Tool\\01_danh_sach_chia_to",
        "metadata_filename": "metadata.json",
        "input_month_year_pattern": "{YYYY}\\{MM}"
    },
    "base_path": "D:\\Linh_Salary_Tool",
    "split_output_path": "D:\\Linh_Salary_Tool\\01_danh_sach_chia_to",
    "_comment": "template_mapping removed (v2.1) - Tab 2 metadata now provides template mapping directly",
    "templates": {
        "gt/bao_ve_template.xlsx": {
            "description": "Bảo vệ template",
            "sheets": {
                "Bang TH nop": {
                    "start_row": 9,
                    "row_step": 2,
                    "columns": {
                        "D": {
                            "source_file": "Danh sách CBCNV bản dùng làm lương T{mm_int}.xlsx",
                            "source_sheet": "Danh sách",
                            "source_range": "$D$1:$AS$1000",
                            "column_index": 42,
                            "description": "Lương cơ bản"
                        }
                    }
                },
                "Lương {mm}": {
                    "start_row": 10,
                    "row_step": 2,
                    "columns": {
                        "Z": {
                            "source_file": "CK tháng {prev_month}.{prev_year}.xlsx",
                            "source_sheet": "Bản gốc T{prev_month}",
                            "source_range": "$C$1:$J$1000",
                            "column_index": 8,
                            "description": "Công tháng trước"
                        },
                        "AA": {
                            "source_file": "CK tháng {prev_month}.{prev_year}.xlsx",
                            "source_sheet": "Bản gốc T{prev_month}",
                            "source_range": "$C$1:$F$1000",
                            "column_index": 4,
                            "description": "Lương tháng trước"
                        }
                    }
                }
            }
        },
        "gt/bep_an_va_cong_vu_template.xlsx": {
            "description": "Bếp ăn & Công vụ template",
            "sheets": {
                "Bang TH nop": {
                    "start_row": 9,
                    "row_step": 2,
                    "columns": {
                        "D": {
                            "source_file": "Danh sách CBCNV bản dùng làm lương T{mm_int}.xlsx",
                            "source_sheet": "Danh sách",
                            "source_range": "$D$1:$AS$1000",
                            "column_index": 42,
                            "description": "Lương cơ bản"
                        },
                        "E": {
                            "source_file": "Danh sách CBCNV làm lương tháng {mm}.xlsx",
                            "source_sheet": "Danh Sách dùng",
                            "source_range": "$E$5:$AG$11000",
                            "column_index": 29,
                            "description": "Ngày vào"
                        },
                        "G": {
                            "source_file": "Chấm công {mm}.{yy_short}.xlsx",
                            "source_sheet": "Công",
                            "source_range": "$C$5:$AO$6000",
                            "column_index": 39,
                            "description": "Công"
                        },
                        "H": {
                            "source_file": "Chấm công {mm}.{yy_short}.xlsx",
                            "source_sheet": "Công",
                            "source_range": "$C$5:$AU$6000",
                            "column_index": 44,
                            "description": "Phép"
                        },
                        "I": {
                            "source_file": "Chấm công {mm}.{yy_short}.xlsx",
                            "source_sheet": "Công",
                            "source_range": "$C$5:$AR$6000",
                            "column_index": 42,
                            "description": "Lễ tết"
                        },
                        "K": {
                            "source_file": "Chấm công {mm}.{yy_short}.xlsx",
                            "source_sheet": "Công",
                            "source_range": "$C$5:$BC$6000",
                            "column_index": 53,
                            "description": "Thai 7 tháng"
                        },
                        "L": {
                            "source_file": "Chấm công {mm}.{yy_short}.xlsx",
                            "source_sheet": "Công",
                            "source_range": "$C$5:$BE$6000",
                            "column_index": 55,
                            "description": "Thêm giờ"
                        },
                        "M": {
                            "source_file": "phụ cấp con nhỏ năm {yyyy}.xlsx",
                            "source_sheet": "Con nhỏ",
                            "source_range": "$A$3:$B$3000",
                            "column_index": 2,
                            "description": "Tiền con nhỏ"
                        },
                        "R": {
                            "source_file": "phụ cấp con nhỏ năm {yyyy}.xlsx",
                            "source_sheet": "Nhóm 6",
                            "source_range": "$C$2:$E$2000",
                            "column_index": 3,
                            "description": "PC ATV"
                        },
                        "T": {
                            "source_file": "phụ cấp con nhỏ năm {yyyy}.xlsx",
                            "source_sheet": "Thâm niên",
                            "source_range": "$B$3:$G$3000",
                            "column_index": 6,
                            "description": "Thâm niên"
                        },
                        "U": {
                            "source_file": "Chấm công {mm}.{yy_short}.xlsx",
                            "source_sheet": "Công",
                            "source_range": "$C$5:$AP$6000",
                            "column_index": 40,
                            "description": "CN"
                        }
                    }
                },
                "Lương {mm}": {
                    "start_row": 10,
                    "row_step": 2,
                    "columns": {
                        "H": {
                            "formula": "=+(F{row}/26/8*{bep_cong_vu_rate})*H{row_minus_1}",
                            "type": "formula",
                            "description": "Tính lương thêm giờ theo tỷ lệ Bếp/Công vụ"
                        },
                        "Y": {
                            "source_file": "CK tháng {prev_month}.{prev_year}.xlsx",
                            "source_sheet": "Bản gốc T{prev_month}",
                            "source_range": "$C$1:$J$1000",
                            "column_index": 8,
                            "description": "Công tháng trước"
                        },
                        "Z": {
                            "source_file": "CK tháng {prev_month}.{prev_year}.xlsx",
                            "source_sheet": "Bản gốc T{prev_month}",
                            "source_range": "$C$1:$F$1000",
                            "column_index": 4,
                            "description": "Lương tháng trước"
                        }
                    }
                }
            }
        },
        "gt/boc_vac_template.xlsx": {
            "description": "Bốc vác template",
            "sheets": {
                "Bang TH nop": {
                    "start_row": 9,
                    "row_step": 2,
                    "columns": {
                        "D": {
                            "source_file": "Danh sách CBCNV bản dùng làm lương T{mm_int}.xlsx",
                            "source_sheet": "Danh sách",
                            "source_range": "$D$1:$AS$1000",
                            "column_index": 42,
                            "description": "Lương cơ bản"
                        },
                        "E": {
                            "source_file": "Danh sách CBCNV làm lương tháng {mm}.xlsx",
                            "source_sheet": "Danh Sách dùng",
                            "source_range": "$E$5:$AG$11000",
                            "column_index": 29,
                            "description": "Ngày vào"
                        },
                        "G": {
                            "source_file": "Chấm công {mm}.{yy_short}.xlsx",
                            "source_sheet": "Công",
                            "source_range": "$C$5:$AO$6000",
                            "column_index": 39,
                            "description": "Công"
                        },
                        "H": {
                            "source_file": "Chấm công {mm}.{yy_short}.xlsx",
                            "source_sheet": "Công",
                            "source_range": "$C$5:$AU$6000",
                            "column_index": 44,
                            "description": "Phép"
                        },
                        "I": {
                            "source_file": "Chấm công {mm}.{yy_short}.xlsx",
                            "source_sheet": "Công",
                            "source_range": "$C$5:$AR$6000",
                            "column_index": 42,
                            "description": "Lễ tết"
                        },
                        "K": {
                            "source_file": "Chấm công {mm}.{yy_short}.xlsx",
                            "source_sheet": "Công",
                            "source_range": "$C$5:$BC$6000",
                            "column_index": 53,
                            "description": "Đi "
                        },
                        "L": {
                            "source_file": "Chấm công {mm}.{yy_short}.xlsx",
                            "source_sheet": "Công",
                            "source_range": "$C$5:$BE$6000",
                            "column_index": 55,
                            "description": "Thêm giờ"
                        },
                        "M": {
                            "source_file": "phụ cấp con nhỏ năm {yyyy}.xlsx",
                            "source_sheet": "Sheet1",
                            "source_range": "$A$3:$B$3000",
                            "column_index": 2,
                            "description": "Tiền con nhỏ"
                        },
                        "Q": {
                            "source_file": "phụ cấp con nhỏ năm {yyyy}.xlsx",
                            "source_sheet": "Nhóm 6",
                            "source_range": "$C$2:$E$2000",
                            "column_index": 3,
                            "description": "PC ATV"
                        },
                        "T": {
                            "source_file": "Chấm công {mm}.{yy_short}.xlsx",
                            "source_sheet": "Công",
                            "source_range": "$C$5:$AP$6000",
                            "column_index": 40,
                            "description": "CN"
                        }
                    }
                },
                "Lương {mm}": {
                    "start_row": 10,
                    "row_step": 2,
                    "columns": {
                        "AC": {
                            "source_file": "CK tháng {prev_month}.{prev_year}.xlsx",
                            "source_sheet": "Bản gốc T{prev_month}",
                            "source_range": "$C$1:$J$1000",
                            "column_index": 8,
                            "description": "Công tháng trước"
                        },
                        "AD": {
                            "source_file": "CK tháng {prev_month}.{prev_year}.xlsx",
                            "source_sheet": "Bản gốc T{prev_month}",
                            "source_range": "$C$1:$F$1000",
                            "column_index": 4,
                            "description": "Lương tháng trước"
                        }
                    }
                }
            }
        },
        "gt/co_dien_template.xlsx": {
            "description": "Cơ điện (Mechanical/Electrical) template",
            "output_file_pattern": "gt/C.điện - {MM}.xlsx",
            "sheets": {
                "Bang TH nop": {
                    "start_row": 9,
                    "row_step": 2,
                    "columns": {
                        "D": {
                            "source_file": "Danh sách CBCNV bản dùng làm lương T{mm_int}.xlsx",
                            "source_sheet": "Danh sách",
                            "source_range": "$D$1:$AS$1000",
                            "column_index": 42,
                            "description": "Lương cơ bản"
                        },
                        "E": {
                            "source_file": "Danh sách CBCNV làm lương tháng {mm}.xlsx",
                            "source_sheet": "Danh Sách dùng",
                            "source_range": "$E$5:$AG$11000",
                            "column_index": 29,
                            "description": "Ngày vào"
                        },
                        "G": {
                            "source_file": "Chấm công {mm}.{yy_short}.xlsx",
                            "source_sheet": "Công",
                            "source_range": "$C$5:$AO$6000",
                            "column_index": 39,
                            "description": "Công"
                        },
                        "H": {
                            "source_file": "Chấm công {mm}.{yy_short}.xlsx",
                            "source_sheet": "Công",
                            "source_range": "$C$5:$AU$6000",
                            "column_index": 44,
                            "description": "Phép"
                        },
                        "I": {
                            "source_file": "Chấm công {mm}.{yy_short}.xlsx",
                            "source_sheet": "Công",
                            "source_range": "$C$5:$AR$6000",
                            "column_index": 42,
                            "description": "Lễ tết"
                        }
                    }
                },
                "Lương {mm}": {
                    "start_row": 10,
                    "row_step": 2,
                    "columns": {
                        "Z": {
                            "source_file": "CK tháng {prev_month}.{prev_year}.xlsx",
                            "source_sheet": "Bản gốc T{prev_month}",
                            "source_range": "$C$1:$J$1000",
                            "column_index": 8,
                            "description": "Công tháng trước"
                        },
                        "AA": {
                            "source_file": "CK tháng {prev_month}.{prev_year}.xlsx",
                            "source_sheet": "Bản gốc T{prev_month}",
                            "source_range": "$C$1:$F$1000",
                            "column_index": 4,
                            "description": "Lương tháng trước"
                        }
                    }
                }
            }
        },
        "gt/ke_hoach_template.xlsx": {
            "description": "Kế hoạch (Planning) template",
            "output_file_pattern": "gt/K.hoạch - {MM}.xlsx",
            "sheets": {
                "Bang TH nop": {
                    "start_row": 9,
                    "row_step": 2,
                    "columns": {
                        "D": {
                            "source_file": "Danh sách CBCNV bản dùng làm lương T{mm_int}.xlsx",
                            "source_sheet": "Danh sách",
                            "source_range": "$D$1:$AS$1000",
                            "column_index": 42,
                            "description": "Lương cơ bản"
                        }
                    }
                },
                "Lương {mm}": {
                    "start_row": 10,
                    "row_step": 2,
                    "columns": {
                        "Z": {
                            "source_file": "CK tháng {prev_month}.{prev_year}.xlsx",
                            "source_sheet": "Bản gốc T{prev_month}",
                            "source_range": "$C$1:$J$1000",
                            "column_index": 8,
                            "description": "Công tháng trước"
                        },
                        "AA": {
                            "source_file": "CK tháng {prev_month}.{prev_year}.xlsx",
                            "source_sheet": "Bản gốc T{prev_month}",
                            "source_range": "$C$1:$F$1000",
                            "column_index": 4,
                            "description": "Lương tháng trước"
                        }
                    }
                }
            }
        },
        "gt/ky_thuat_template.xlsx": {
            "description": "Kỹ thuật (Technical) template",
            "output_file_pattern": "gt/K.thuật - {MM}.xlsx",
            "sheets": {
                "Bang TH nop": {
                    "start_row": 9,
                    "row_step": 2,
                    "columns": {
                        "D": {
                            "source_file": "Danh sách CBCNV bản dùng làm lương T{mm_int}.xlsx",
                            "source_sheet": "Danh sách",
                            "source_range": "$D$1:$AS$1000",
                            "column_index": 42,
                            "description": "Lương cơ bản"
                        }
                    }
                },
                "Lương {mm}": {
                    "start_row": 10,
                    "row_step": 2,
                    "columns": {
                        "Z": {
                            "source_file": "CK tháng {prev_month}.{prev_year}.xlsx",
                            "source_sheet": "Bản gốc T{prev_month}",
                            "source_range": "$C$1:$J$1000",
                            "column_index": 8,
                            "description": "Công tháng trước"
                        },
                        "AA": {
                            "source_file": "CK tháng {prev_month}.{prev_year}.xlsx",
                            "source_sheet": "Bản gốc T{prev_month}",
                            "source_range": "$C$1:$F$1000",
                            "column_index": 4,
                            "description": "Lương tháng trước"
                        }
                    }
                }
            }
        },
        "gt/van_phong_template.xlsx": {
            "description": "Văn phòng (Office) template",
            "output_file_pattern": "gt/VP - {MM}.xlsx",
            "sheets": {
                "Bang TH nop": {
                    "start_row": 9,
                    "row_step": 2,
                    "columns": {
                        "D": {
                            "source_file": "Danh sách CBCNV bản dùng làm lương T{mm_int}.xlsx",
                            "source_sheet": "Danh sách",
                            "source_range": "$D$1:$AS$1000",
                            "column_index": 42,
                            "description": "Lương cơ bản"
                        }
                    }
                },
                "Lương {mm}": {
                    "start_row": 10,
                    "row_step": 2,
                    "columns": {
                        "Z": {
                            "source_file": "CK tháng {prev_month}.{prev_year}.xlsx",
                            "source_sheet": "Bản gốc T{prev_month}",
                            "source_range": "$C$1:$J$1000",
                            "column_index": 8,
                            "description": "Công tháng trước"
                        },
                        "AA": {
                            "source_file": "CK tháng {prev_month}.{prev_year}.xlsx",
                            "source_sheet": "Bản gốc T{prev_month}",
                            "source_range": "$C$1:$F$1000",
                            "column_index": 4,
                            "description": "Lương tháng trước"
                        }
                    }
                }
            }
        },
        "sx/to_may_template.xlsx": {
            "description": "Tổ may (Sewing teams: Tổ 1-12, Tổ trưởng, Thời vụ, Cắt, H.thiện, etc) template",
            "output_file_pattern": "sx/(Tổ|Tổ trưởng|Thời vụ|Cắt|H.thiện|Điều động|Hậu giặt|In) .* - {MM}.xlsx",
            "sheets": {
                "Danh sách": {
                    "start_row": 5,
                    "row_step": 1,
                    "columns": {}
                },
                "Lương {mm}": {
                    "start_row": 10,
                    "row_step": 2,
                    "columns": {
                        "Z": {
                            "source_file": "CK tháng {prev_month}.{prev_year}.xlsx",
                            "source_sheet": "Bản gốc T{prev_month}",
                            "source_range": "$C$1:$J$1000",
                            "column_index": 8,
                            "description": "Công tháng trước"
                        },
                        "AA": {
                            "source_file": "CK tháng {prev_month}.{prev_year}.xlsx",
                            "source_sheet": "Bản gốc T{prev_month}",
                            "source_range": "$C$1:$F$1000",
                            "column_index": 4,
                            "description": "Lương tháng trước"
                        }
                    }
                }
            }
        },
        "sx/kiem_hoa_template.xlsx": {
            "description": "Kiểm hóa (QC/Inspection) template",
            "output_file_pattern": "sx/Kiểm hóa - {MM}.xlsx",
            "sheets": {
                "Danh sách": {
                    "start_row": 5,
                    "row_step": 1,
                    "columns": {}
                },
                "Lương {mm}": {
                    "start_row": 10,
                    "row_step": 2,
                    "columns": {
                        "Z": {
                            "source_file": "CK tháng {prev_month}.{prev_year}.xlsx",
                            "source_sheet": "Bản gốc T{prev_month}",
                            "source_range": "$C$1:$J$1000",
                            "column_index": 8,
                            "description": "Công tháng trước"
                        },
                        "AA": {
                            "source_file": "CK tháng {prev_month}.{prev_year}.xlsx",
                            "source_sheet": "Bản gốc T{prev_month}",
                            "source_range": "$C$1:$F$1000",
                            "column_index": 4,
                            "description": "Lương tháng trước"
                        }
                    }
                }
            }
        }
    }
}

salary_thread.py:
import os
import openpyxl
from typing import Dict, Optional, Tuple
from logic.salary_config_parser import get_salary_config
from logic.metadata_manager import get_metadata_manager


def deploy_salary_from_metadata(year: int, month: int, log_callback, progress_callback=None) -> Tuple[bool, str]:
    """
    ✨ NEW: Triển khai công lương từ metadata - Tuần tự + Tối ưu hóa
    
    Workflow:
    1. Đọc metadata.json từ D:\Linh_Salary_Tool\01_danh_sach_chia_to\YYYY\MM\data\metadata.json
    2. Từ metadata xác định: file output, template tương ứng
    3. Tuần tự xử lý từng file:
       - Load file Excel
       - Áp dụng công thức từ config
       - Linking dữ liệu từ source files
       - Lưu file
    4. Tối ưu hóa: Cache file sources, batch operations
    
    Args:
        year: Năm (VD: 2026)
        month: Tháng (VD: 4)
        log_callback: Callback để ghi log
        progress_callback: Callback để cập nhật tiến độ (0-100)
        
    Returns:
        (success: bool, message: str)
    """
    try:
        config = get_salary_config()
    except Exception as e:
        return False, f"❌ Lỗi đọc config JSON: {str(e)}"
    
    # Bước 1: Tải metadata
    metadata_mgr = get_metadata_manager()
    if not metadata_mgr.load_metadata(year, month):
        return False, f"❌ Không tìm thấy metadata.json tại: {metadata_mgr.metadata_path}"
    
    # Kiểm tra tính hợp lệ metadata
    is_valid, errors = metadata_mgr.validate_metadata(log_callback)
    if not is_valid:
        return False, "❌ Metadata không hợp lệ:\n" + "\n".join(errors)
    
    log_callback(f"📊 TRIỂN KHAI CÔNG LƯƠNG - Từ Metadata")
    log_callback(f"📅 Năm/Tháng: {year}/{month:02d}")
    log_callback(f"📂 Metadata: {metadata_mgr.metadata_path}\n")
    
    # Bước 2: Lấy danh sách file và template từ metadata
    all_groups = metadata_mgr.get_all_groups_with_templates()
    if not all_groups:
        return False, "❌ Metadata không chứa thông tin file/template nào"
    
    log_callback(f"🔍 Tìm thấy {len(all_groups)} group nhân viên\n")
    
    # Bước 3: Tuần tự xử lý từng file
    processed = 0
    warnings = []
    
    # Cache để tối ưu hiệu năng: lưu file sources đã resolve
    file_source_cache: Dict[str, Optional[str]] = {}
    
    for idx, (group_name, template_key, output_filename) in enumerate(all_groups):
        try:
            # Xây dựng đường dẫn đầy đủ file output
            split_output = config.config.get("split_output_path", "D:\\Linh_Salary_Tool\\01_danh_sach_chia_to")
            mm = f"{int(month):02d}"
            yyyy = str(year)
            excel_file = os.path.join(split_output, yyyy, mm, output_filename)
            
            if not os.path.exists(excel_file):
                warning = f"⚠️ [{idx + 1}/{len(all_groups)}] File không tìm thấy: {output_filename}"
                log_callback(warning)
                warnings.append(warning)
                continue
            
            log_callback(f"📄 [{idx + 1}/{len(all_groups)}] Xử lý: {group_name}")
            log_callback(f"   📋 Template: {template_key}")
            log_callback(f"   📁 File: {output_filename}\n")
            
            # Xử lý file này
            success = _process_excel_file(
                excel_file,
                template_key,
                year,
                month,
                config,
                log_callback,
                file_source_cache
            )
            
            if success:
                processed += 1
            else:
                warnings.append(f"⚠️ File {output_filename}: Xử lý có lỗi")
            
        except Exception as e:
            warning = f"❌ [{idx + 1}/{len(all_groups)}] Lỗi xử lý group '{group_name}': {str(e)}"
            log_callback(warning)
            warnings.append(warning)
            import traceback
            log_callback(f"   {traceback.format_exc()}\n")
        
        # Cập nhật tiến độ
        if progress_callback:
            progress_callback(int(((idx + 1) / len(all_groups)) * 100))
    
    if progress_callback:
        progress_callback(100)
    
    # Kết quả
    result_msg = f"\n✅ Hoàn tất! Đã triển khai {processed}/{len(all_groups)} file."
    if warnings:
        result_msg += f"\n⚠️ Có {len(warnings)} cảnh báo - Kiểm tra log"
    
    return True, result_msg


def _process_excel_file(
    excel_file: str,
    template_key: str,
    year: int,
    month: int,
    config,
    log_callback,
    file_source_cache: Dict[str, Optional[str]]
) -> bool:
    """
    Xử lý một file Excel duy nhất.
    Tối ưu: Sử dụng cache cho file sources.
    
    Returns:
        True nếu thành công, False nếu lỗi
    """
    yyyy = int(year)
    mm = int(month)
    yyyy_str = str(yyyy)
    mm_str = f"{mm:02d}"
    mm_int = mm
    
    # Tính tháng trước
    if mm_int == 1:
        prev_month_int = 12
        prev_year_int = yyyy - 1
    else:
        prev_month_int = mm_int - 1
        prev_year_int = yyyy
    
    prev_month = f"{prev_month_int:02d}"
    prev_year = str(prev_year_int)
    yy_short = yyyy_str[-2:]
    prev_yy_short = prev_year[-2:]
    
    # Validate template
    is_valid, errors = config.validate_template(template_key)
    if not is_valid:
        log_callback(f"   ❌ Template config lỗi: {errors[0]}")
        return False
    
    try:
        wb = openpyxl.load_workbook(excel_file)
    except Exception as e:
        log_callback(f"   ❌ Lỗi mở file Excel: {str(e)}")
        return False
    
    file_modified = False
    template_config = config.get_template_config(template_key)
    
    try:
        # Xác định tỷ lệ cho Bếp/Công vụ (dựa trên tên file)
        file_lower = os.path.basename(excel_file).lower()
        bep_cong_vu_rate = "100%"
        if "bếp" in file_lower or "bep" in file_lower:
            bep_cong_vu_rate = "125%"
        elif "công vụ" in file_lower or "cong vu" in file_lower:
            bep_cong_vu_rate = "115%"
        
        # Lặp qua tất cả sheets trong template config
        for sheet_pattern, sheet_cfg in template_config.get("sheets", {}).items():
            # Resolve dynamic sheet name
            resolved_sheet_name = sheet_pattern.replace("{mm}", mm_str)
            
            # Tìm sheet trong workbook (case-insensitive)
            ws = None
            actual_sheet_name = None
            for name in wb.sheetnames:
                if name.lower() == resolved_sheet_name.lower():
                    ws = wb[name]
                    actual_sheet_name = name
                    break
            
            if not ws:
                log_callback(f"   ℹ️ Sheet '{resolved_sheet_name}' không tìm thấy → Bỏ qua")
                continue
            
            # Mở khóa sheet nếu có bảo vệ
            try:
                ws.protection.sheet = False
            except Exception:
                pass
            
            start_row = sheet_cfg.get("start_row", 9)
            row_step = sheet_cfg.get("row_step", 2)
            
            # Tính số lượng nhân viên từ dữ liệu trong cột C (mã nhân viên)
            num_employees = 0
            for row in range(start_row, start_row + 500, row_step):
                cell_value = ws.cell(row=row, column=3).value
                if cell_value and str(cell_value).strip():
                    num_employees += 1
                else:
                    break
            
            if num_employees == 0:
                log_callback(f"   ℹ️ Sheet '{actual_sheet_name}': Không có dữ liệu nhân viên")
                continue
            
            log_callback(f"   📋 Sheet: {actual_sheet_name} ({num_employees} nhân viên)")
            
            # Lặp qua từng cột có config
            columns_config = sheet_cfg.get("columns", {})
            for col_letter, col_cfg in columns_config.items():
                parsed = config.parse_column_config(col_cfg)
                col_num = ord(col_letter.upper()) - ord('A') + 1
                
                if parsed["type"] == "formula":
                    # Công thức tĩnh - áp dụng trực tiếp
                    formula_template = parsed["formula"]
                    
                    for idx in range(num_employees):
                        current_row = start_row + (idx * row_step)
                        formula = formula_template.format(
                            row=current_row,
                            row_minus_1=current_row - 1,
                            row_plus_1=current_row + 1,
                            bep_cong_vu_rate=bep_cong_vu_rate
                        )
                        try:
                            ws.cell(row=current_row, column=col_num).value = formula
                        except Exception as e:
                            log_callback(f"      ❌ Cột {col_letter}, dòng {current_row}: {str(e)}")
                    file_modified = True
                
                elif parsed["type"] == "source":
                    # Linking từ file ngoài - Sử dụng CACHE
                    source_file = parsed["file"]
                    source_sheet = parsed["sheet"]
                    source_range = parsed["range"]
                    col_index = parsed.get("column_index", 1)
                    
                    # Resolve file source (với cache)
                    cache_key = f"{source_file}:{yyyy}:{mm}"
                    if cache_key not in file_source_cache:
                        file_source_cache[cache_key] = config.resolve_file_path(
                            source_file, yyyy, mm, lambda x: None  # Silent log
                        )
                    
                    source_path = file_source_cache[cache_key]
                    
                    if source_path is None:
                        # File không tồn tại → để trống + warning
                        log_callback(f"      ⚠️ Cột {col_letter}: File '{source_file}' không tìm thấy")
                        
                        for idx in range(num_employees):
                            current_row = start_row + (idx * row_step)
                            ws.cell(row=current_row, column=col_num).value = None
                        file_modified = True
                    else:
                        # File tồn tại → tạo formula linking
                        for idx in range(num_employees):
                            current_row = start_row + (idx * row_step)
                            # Excel external link formula
                            formula = f"=IFERROR(VLOOKUP(C{current_row},'{source_path}[{os.path.basename(source_path)}]{source_sheet}'!{source_range},{col_index},0),\"\")"
                            try:
                                ws.cell(row=current_row, column=col_num).value = formula
                            except Exception as e:
                                log_callback(f"      ❌ Cột {col_letter}, dòng {current_row}: {str(e)}")
                        file_modified = True
        
        # Lưu file nếu có thay đổi
        if file_modified:
            wb.save(excel_file)
            log_callback(f"   ✅ Đã lưu file\n")
        else:
            log_callback(f"   ℹ️ Không có thay đổi\n")
        
        return True
    
    except Exception as e:
        log_callback(f"   ❌ Lỗi xử lý: {str(e)}")
        import traceback
        log_callback(f"   {traceback.format_exc()}")
        return False
    
    finally:
        wb.close()

salary_widget.py:
from datetime import datetime
from PySide6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QLabel,
    QPushButton, QComboBox, QTextEdit, QMessageBox, QFrame, QProgressBar
)
from PySide6.QtCore import Qt
from PySide6.QtGui import QFont
from ui.tab5_salary_deploy.salary_thread import SalaryDeployThread


class SalaryDeployWidget(QWidget):
    """
    Giao diện Tab 5: Triển khai Công Lương (Salary Deployment).
    
    ✨ Version 2.1 - Metadata-Based Auto-Detection:
    - Tự động đọc metadata.json từ Tab 2 output
    - Xác định template cho từng file
    - Triển khai tuần tự, tối ưu hiệu năng
    - Chỉ cần nhập tháng/năm, tất cả khác tự động!
    
    Workflow:
    1. User chọn tháng/năm
    2. Nhấn "🚀 TRIỂN KHAI NGAY"
    3. Hệ thống:
       - Đọc metadata.json từ D:\Linh_Salary_Tool\01_danh_sach_chia_to\YYYY\MM\data\metadata.json
       - Xác định template cho từng file từ metadata
       - Tuần tự fill công thức vào các sheet tương ứng
       - Linking dữ liệu từ source files
    4. Hiển thị kết quả
    """
    def __init__(self):
        super().__init__()
        self.setObjectName("SalaryWidget")
        self.thread = None
        self.setStyleSheet("""
            QWidget#SalaryWidget { background-color: white; border-radius: 8px; }
            QLabel { font-size: 13px; font-weight: bold; color: #374151; }
            QComboBox { 
                padding: 8px; border: 1px solid #D1D5DB; border-radius: 5px; background-color: white; color: #111827; font-size: 14px;
            }
            QPushButton {
                background-color: #3B82F6; color: white; border: none; border-radius: 6px; padding: 8px 15px; font-size: 13px; font-weight: bold;
            }
            QPushButton:hover { background-color: #2563EB; }
            QPushButton:disabled { background-color: #9CA3AF; }
            QPushButton#ProcessBtn { background-color: #10B981; font-size: 16px; padding: 15px 20px; }
            QPushButton#ProcessBtn:hover { background-color: #059669; }
            QProgressBar {
                border: 1px solid #E5E7EB; border-radius: 5px; text-align: center; background-color: #F3F4F6; color: #111827; font-weight: bold; height: 22px;
            }
            QProgressBar::chunk { background-color: #10B981; border-radius: 4px; }
            QTextEdit {
                background-color: #1E1E1E; color: #10B981; font-family: Consolas, monospace; font-size: 13px; border: 1px solid #374151; border-radius: 6px; padding: 10px;
            }
            QFrame#SectionFrame { border: 1px solid #E5E7EB; border-radius: 8px; background-color: #F9FAFB; padding: 15px; }
        """)
        self.init_ui()

    def init_ui(self):
        layout = QVBoxLayout(self)
        layout.setContentsMargins(20, 20, 20, 20)
        layout.setSpacing(15)

        # Tiêu đề
        title = QLabel("💰 TRIỂN KHAI CÔNG THỨC LƯƠNG")
        title.setFont(QFont("Arial", 18, QFont.Bold))
        title.setAlignment(Qt.AlignCenter)
        title.setStyleSheet("color: #1F2937; margin-bottom: 5px;")
        layout.addWidget(title)

        # Mô tả
        desc = QLabel("Chỉ cần chọn tháng/năm, hệ thống tự động:\n"
                     "✅ Đọc metadata từ Tab 2\n"
                     "✅ Xác định template cho từng file\n"
                     "✅ Fill công thức & linking dữ liệu")
        desc.setAlignment(Qt.AlignCenter)
        desc.setStyleSheet("color: #666; font-size: 12px; margin-bottom: 10px;")
        layout.addWidget(desc)

        # Phần nhập Tháng/Năm
        date_frame = QFrame()
        date_frame.setObjectName("SectionFrame")
        date_layout = QHBoxLayout(date_frame)
        
        now = datetime.now()
        
        date_layout.addWidget(QLabel("📅 Tháng:"))
        self.combo_month = QComboBox()
        self.combo_month.addItems([f"{i:02d}" for i in range(1, 13)])
        self.combo_month.setCurrentText(f"{now.month:02d}")
        self.combo_month.setMinimumWidth(80)
        date_layout.addWidget(self.combo_month)
        
        date_layout.addWidget(QLabel("Năm:"))
        self.combo_year = QComboBox()
        self.combo_year.addItems([str(y) for y in range(2024, 2031)])
        self.combo_year.setCurrentText(str(now.year))
        self.combo_year.setMinimumWidth(100)
        date_layout.addWidget(self.combo_year)
        
        date_layout.addStretch()
        layout.addWidget(date_frame)

        # Nút thực hiện
        self.btn_run = QPushButton("🚀 TRIỂN KHAI NGAY")
        self.btn_run.setObjectName("ProcessBtn")
        self.btn_run.clicked.connect(self.start_processing)
        self.btn_run.setCursor(Qt.PointingHandCursor)
        layout.addWidget(self.btn_run)
        
        # Progress bar
        self.progress_bar = QProgressBar()
        self.progress_bar.setVisible(False)
        layout.addWidget(self.progress_bar)

        # Nhật ký hoạt động
        log_label = QLabel("📝 Nhật Ký Hoạt Động:")
        log_label.setStyleSheet("color: #374151; font-weight: bold; font-size: 13px; margin-top: 10px;")
        layout.addWidget(log_label)
        
        self.log_area = QTextEdit()
        self.log_area.setReadOnly(True)
        layout.addWidget(self.log_area, stretch=1)

    def log_callback(self, message: str):
        """Append message to log area"""
        self.log_area.append(message)

    def progress_callback(self, value: int):
        """Update progress bar"""
        self.progress_bar.setValue(value)

    def start_processing(self):
        """
        Bắt đầu triển khai công lương từ metadata.
        """
        month = int(self.combo_month.currentText())
        year = int(self.combo_year.currentText())

        self.btn_run.setEnabled(False)
        self.log_area.clear()
        self.progress_bar.setValue(0)
        self.progress_bar.setVisible(True)

        # Khởi tạo thread xử lý
        self.thread = SalaryDeployThread(year, month)
        self.thread.log_signal.connect(self.log_callback)
        self.thread.progress_signal.connect(self.progress_callback)
        self.thread.finished_signal.connect(self.on_deployment_finished)
        self.thread.start()

    def on_deployment_finished(self, success: bool, message: str):
        """
        Xử lý kết thúc deployment.
        """
        if success:
            QMessageBox.information(
                self, "✅ Thành Công",
                f"Triển khai công lương thành công!\n\n{message}"
            )
        else:
            QMessageBox.critical(
                self, "❌ Lỗi",
                f"Lỗi triển khai:\n\n{message}"
            )
        
        self.btn_run.setEnabled(True)
        self.progress_bar.setVisible(False)

D:\Linh_Salary_Tool\01_danh_sach_chia_to\2026\04\data\metadata.json:
{
  "year_month": "2026-04",
  "output_dir": "D:/Linh_Salary_Tool\\01_danh_sach_chia_to\\2026\\04",
  "mapping": {
    "Danh sách chưa được xử lý - 04.xlsx": {
      "template": "chua_xu_ly_template.xlsx",
      "group_name": "Chưa xử lý",
      "employee_count": 33,
      "folder": "root"
    },
    "gt\\K.hoạch - 04.xlsx": {
      "template": "gt/ke_hoach_template.xlsx",
      "group_name": "K.hoạch",
      "employee_count": 13,
      "folder": "gt"
    },
    "sx\\Kiểm hóa - 04.xlsx": {
      "template": "sx/kiem_hoa_template.xlsx",
      "group_name": "Kiểm hóa",
      "employee_count": 17,
      "folder": "sx"
    },
    "sx\\In - 04.xlsx": {
      "template": "sx/to_may_template.xlsx",
      "group_name": "In",
      "employee_count": 18,
      "folder": "sx"
    },
    "gt\\C.điện - 04.xlsx": {
      "template": "gt/co_dien_template.xlsx",
      "group_name": "C.điện",
      "employee_count": 9,
      "folder": "gt"
    },
    "sx\\Tổ 4 - 04.xlsx": {
      "template": "sx/to_may_template.xlsx",
      "group_name": "Tổ 4",
      "employee_count": 17,
      "folder": "sx"
    },
    "sx\\Tổ 10 - 04.xlsx": {
      "template": "sx/to_may_template.xlsx",
      "group_name": "Tổ 10",
      "employee_count": 17,
      "folder": "sx"
    },
    "sx\\Điều động - 04.xlsx": {
      "template": "sx/to_may_template.xlsx",
      "group_name": "Điều động",
      "employee_count": 12,
      "folder": "sx"
    },
    "sx\\Tổ trưởng may - 04.xlsx": {
      "template": "sx/to_may_template.xlsx",
      "group_name": "Tổ trưởng may",
      "employee_count": 12,
      "folder": "sx"
    },
    "sx\\Cắt - 04.xlsx": {
      "template": "sx/to_may_template.xlsx",
      "group_name": "Cắt",
      "employee_count": 24,
      "folder": "sx"
    },
    "sx\\Tổ 12 - 04.xlsx": {
      "template": "sx/to_may_template.xlsx",
      "group_name": "Tổ 12",
      "employee_count": 9,
      "folder": "sx"
    },
    "sx\\Tổ 1 - 04.xlsx": {
      "template": "sx/to_may_template.xlsx",
      "group_name": "Tổ 1",
      "employee_count": 20,
      "folder": "sx"
    },
    "sx\\Tổ 2 - 04.xlsx": {
      "template": "sx/to_may_template.xlsx",
      "group_name": "Tổ 2",
      "employee_count": 18,
      "folder": "sx"
    },
    "sx\\Tổ 8 - 04.xlsx": {
      "template": "sx/to_may_template.xlsx",
      "group_name": "Tổ 8",
      "employee_count": 16,
      "folder": "sx"
    },
    "sx\\Tổ 11 - 04.xlsx": {
      "template": "sx/to_may_template.xlsx",
      "group_name": "Tổ 11",
      "employee_count": 17,
      "folder": "sx"
    },
    "sx\\H.thiện - 04.xlsx": {
      "template": "sx/to_may_template.xlsx",
      "group_name": "H.thiện",
      "employee_count": 17,
      "folder": "sx"
    },
    "sx\\Tổ 5 - 04.xlsx": {
      "template": "sx/to_may_template.xlsx",
      "group_name": "Tổ 5",
      "employee_count": 19,
      "folder": "sx"
    },
    "gt\\K.thuật - 04.xlsx": {
      "template": "gt/ky_thuat_template.xlsx",
      "group_name": "K.thuật",
      "employee_count": 13,
      "folder": "gt"
    },
    "gt\\Bảo vệ - 04.xlsx": {
      "template": "gt/bao_ve_template.xlsx",
      "group_name": "Bảo vệ",
      "employee_count": 7,
      "folder": "gt"
    },
    "sx\\Tổ 3 - 04.xlsx": {
      "template": "sx/to_may_template.xlsx",
      "group_name": "Tổ 3",
      "employee_count": 16,
      "folder": "sx"
    },
    "sx\\Tổ 7 - 04.xlsx": {
      "template": "sx/to_may_template.xlsx",
      "group_name": "Tổ 7",
      "employee_count": 16,
      "folder": "sx"
    },
    "sx\\Tổ 6 - 04.xlsx": {
      "template": "sx/to_may_template.xlsx",
      "group_name": "Tổ 6",
      "employee_count": 13,
      "folder": "sx"
    },
    "gt\\Công vụ - 04.xlsx": {
      "template": "gt/bep_an_va_cong_vu_template.xlsx",
      "group_name": "Công vụ",
      "employee_count": 4,
      "folder": "gt"
    },
    "gt\\VP - 04.xlsx": {
      "template": "gt/van_phong_template.xlsx",
      "group_name": "VP",
      "employee_count": 5,
      "folder": "gt"
    },
    "gt\\Bếp ăn - 04.xlsx": {
      "template": "gt/bep_an_va_cong_vu_template.xlsx",
      "group_name": "Bếp ăn",
      "employee_count": 4,
      "folder": "gt"
    },
    "sx\\Hậu giặt - 04.xlsx": {
      "template": "sx/to_may_template.xlsx",
      "group_name": "Hậu giặt",
      "employee_count": 4,
      "folder": "sx"
    },
    "sx\\Tổ 9 - 04.xlsx": {
      "template": "sx/to_may_template.xlsx",
      "group_name": "Tổ 9",
      "employee_count": 15,
      "folder": "sx"
    },
    "gt\\Bốc vác - 04.xlsx": {
      "template": "gt/boc_vac_template.xlsx",
      "group_name": "Bốc vác",
      "employee_count": 4,
      "folder": "gt"
    },
    "sx\\Thời vụ - 04.xlsx": {
      "template": "sx/to_may_template.xlsx",
      "group_name": "Thời vụ",
      "employee_count": 19,
      "folder": "sx"
    }
  }
}