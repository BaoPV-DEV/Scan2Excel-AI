import os
import json
from typing import Dict, List, Optional, Tuple


class MetadataManager:
    """
    QUẢN LÝ METADATA TAB 2 → TAB 5 PIPELINE

    Vai trò:
    - Load metadata.json
    - Mapping file output → template
    - Resolve path an toàn (EXE-safe)
    """

    def __init__(self, metadata_path: str = None):
        self.metadata_path = metadata_path
        self.metadata = {}

    # =========================================================
    # LOAD METADATA
    # =========================================================
    def load_metadata(self, year: int, month: int) -> bool:
        """
        Load metadata.json từ path cấu hình hoặc default
        """

        if self.metadata_path is None:
            from app.utils.paths import get_metadata_path
            self.metadata_path = get_metadata_path(year, month)

        if not self.metadata_path or not os.path.exists(self.metadata_path):
            return False

        try:
            with open(self.metadata_path, "r", encoding="utf-8") as f:
                self.metadata = json.load(f)
            return True

        except Exception as e:
            print(f"❌ Metadata load error: {e}")
            self.metadata = {}
            return False

    # =========================================================
    # YEAR-MONTH
    # =========================================================
    def get_year_month(self) -> Optional[str]:
        return self.metadata.get("year_month")

    # =========================================================
    # GET MAPPING (TAB 2 OUTPUT)
    # =========================================================
    def get_split_files(self) -> Dict[str, Dict]:
        """
        mapping format:
        {
            "sx/Tổ 1 - 04.xlsx": {
                "template": "...",
                "employee_count": 10
            }
        }
        """
        return self.metadata.get("mapping", {})

    # =========================================================
    # SAFE PATH RESOLVER (IMPORTANT FOR EXE)
    # =========================================================
    def _normalize_drive_path(self, path: str) -> str:
        """
        Fix case lỗi D:\ → C:\ khi chạy EXE khác môi trường
        """
        if not path:
            return path

        if not os.path.exists("D:/") and path.startswith("D:\\"):
            return "C:\\" + path[3:]

        return path

    # =========================================================
    # RESOLVE OUTPUT FILE PATH
    # =========================================================
    def get_output_file_path(self, year: int, month: int, relative_path: str) -> Optional[str]:
        """
        Trả về full path file output từ metadata
        """

        from app.logic.salary_config_parser import get_salary_config
        from app.utils.paths import get_base_path

        config = get_salary_config()

        default_split = os.path.join(
            get_base_path(),
            "01_danh_sach_chia_to"
        )

        split_output = config.config.get(
            "split_output_path",
            default_split
        )

        split_output = self._normalize_drive_path(split_output)

        mm = f"{int(month):02d}"
        yyyy = str(year)

        output_dir = self.metadata.get("output_dir")

        if output_dir:
            output_dir = self._normalize_drive_path(output_dir)
        else:
            output_dir = os.path.join(split_output, yyyy, mm)

        full_path = os.path.join(output_dir, relative_path)

        return full_path if os.path.exists(full_path) else None

    # =========================================================
    # GET TEMPLATE FOR FILE
    # =========================================================
    def get_template_for_group(self, relative_path: str) -> Optional[str]:
        mapping = self.get_split_files()
        return mapping.get(relative_path, {}).get("template")

    # =========================================================
    # GET ALL GROUPS
    # =========================================================
    def get_all_groups_with_templates(self) -> List[Tuple[str, str, str]]:
        """
        return:
        [(file, template, folder)]
        """

        result = []
        mapping = self.get_split_files()

        for relative_path, info in mapping.items():
            template = info.get("template")
            folder = info.get("folder", "root")

            if template:
                result.append((relative_path, template, folder))

        return result

    # =========================================================
    # VALIDATION
    # =========================================================
    def validate_metadata(self, log_callback=None) -> Tuple[bool, List[str]]:
        """
        Check metadata integrity
        """

        errors = []

        if not self.metadata:
            errors.append("Metadata chưa được load")
            return False, errors

        if "year_month" not in self.metadata:
            errors.append("Missing year_month")

        mapping = self.metadata.get("mapping")

        if not mapping:
            errors.append("Missing mapping")
        else:
            for file_path, info in mapping.items():
                if not isinstance(info, dict):
                    errors.append(f"Invalid mapping format: {file_path}")
                    continue

                if "template" not in info:
                    errors.append(f"Missing template: {file_path}")

        if log_callback:
            for e in errors:
                log_callback(f"❌ {e}")

        return len(errors) == 0, errors


# =========================================================
# SINGLETON PATTERN (GLOBAL USAGE)
# =========================================================
_global_metadata = None


def get_metadata_manager(metadata_path: str = None) -> MetadataManager:
    global _global_metadata

    if _global_metadata is None:
        _global_metadata = MetadataManager(metadata_path)

    return _global_metadata