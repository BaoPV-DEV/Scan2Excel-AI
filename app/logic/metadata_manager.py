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
        Tải metadata.json từ đường dẫn tập trung: 
        D:\Linh_Salary_Tool\04_data\{YYYY}\{MM}\metadata.json
        
        Args:
            year: Năm (VD: 2026)
            month: Tháng (VD: 4)
            
        Returns:
            True nếu thành công, False nếu không
        """
        if self.metadata_path is None:
            from app.utils.paths import get_metadata_path
            self.metadata_path = get_metadata_path(year, month)
        
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
        from app.logic.salary_config_parser import get_salary_config
        from app.utils.paths import get_base_path
        
        config = get_salary_config()
        default_split = os.path.join(get_base_path(), "01_danh_sach_chia_to")
        split_output = config.config.get("split_output_path", default_split)
        
        # Nếu D không tồn tại mà split_output bắt đầu bằng D:\, tự chuyển sang C:\
        if not os.path.exists("D:/") and split_output.startswith("D:\\"):
            split_output = "C:\\" + split_output[3:]
            
        mm = f"{int(month):02d}"
        yyyy = str(year)
        
        # Lấy output_dir từ metadata
        output_dir = self.metadata.get("output_dir")
        if output_dir and not os.path.exists("D:/") and output_dir.startswith("D:\\"):
            output_dir = "C:\\" + output_dir[3:]
            
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

