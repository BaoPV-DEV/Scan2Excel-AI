import os
import json
from typing import Dict, List, Optional, Tuple


class SalarySourceConfig:
    """
    TEMPLATE CONFIG ENGINE

    Vai trò:
    - Load salary_sources.json
    - Resolve file theo year/month
    - Parse column logic (formula / team_formula / source)
    """

    def __init__(self, config_path: str = None):
        from app.utils.paths import get_salary_sources_path, get_base_path

        self.config_path = config_path or get_salary_sources_path()
        self.config = self._load_config()

        self.base_path = self.config.get("base_path", get_base_path())
        self.templates = self.config.get("templates", {})

    # =========================================================
    # LOAD CONFIG SAFE
    # =========================================================
    def _load_config(self) -> dict:
        try:
            if not os.path.exists(self.config_path):
                raise FileNotFoundError(self.config_path)

            with open(self.config_path, "r", encoding="utf-8") as f:
                return json.load(f)

        except json.JSONDecodeError as e:
            raise ValueError(f"JSON invalid: {e}")

        except Exception as e:
            raise RuntimeError(f"Config load failed: {e}")

    # =========================================================
    # TEMPLATE ACCESS
    # =========================================================
    def get_template_config(self, template_key: str) -> Optional[Dict]:
        return self.templates.get(template_key)

    # =========================================================
    # PATH NORMALIZATION (IMPORTANT FOR EXE)
    # =========================================================
    def _normalize_drive(self, path: str) -> str:
        """
        Fix case D:\ → C:\ when running EXE on different machine
        """
        if not path:
            return path

        if not os.path.exists("D:/") and path.startswith("D:\\"):
            return "C:\\" + path[3:]

        return path

    # =========================================================
    # RESOLVE FILE PATH
    # =========================================================
    def resolve_file_path(
        self,
        file_pattern: str,
        year: int,
        month: int,
        log_callback=None
    ) -> Optional[str]:

        yyyy = str(year)
        mm = f"{int(month):02d}"
        yy_short = yyyy[-2:]
        mm_int = int(month)

        # previous month logic
        if mm_int == 1:
            prev_month = 12
            prev_year = year - 1
        else:
            prev_month = mm_int - 1
            prev_year = year

        resolved = file_pattern.format(
            yyyy=yyyy,
            yy_short=yy_short,
            mm=mm,
            mm_int=mm_int,
            prev_month=f"{prev_month:02d}",
            prev_year=str(prev_year),
            prev_yy_short=str(prev_year)[-2:]
        )

        # build path
        full_path = os.path.join(
            self.base_path,
            str(year),
            mm,
            resolved
        )

        full_path = self._normalize_drive(full_path)

        if os.path.exists(full_path):
            return full_path

        if log_callback:
            log_callback(f"⚠️ Missing file: {full_path}")

        return None

    # =========================================================
    # SHEET CONFIG
    # =========================================================
    def get_sheet_config(self, template_key: str, sheet_name: str) -> Optional[Dict]:
        template = self.get_template_config(template_key)
        if not template:
            return None

        sheets = template.get("sheets", {})

        for pattern, cfg in sheets.items():
            if sheet_name == pattern or sheet_name.lower().startswith(pattern.lower()):
                return cfg

        return None

    # =========================================================
    # COLUMN SOURCES
    # =========================================================
    def get_column_sources(self, template_key: str, sheet_name: str) -> Dict:
        sheet = self.get_sheet_config(template_key, sheet_name)
        if not sheet:
            return {}

        return sheet.get("columns", {})

    # =========================================================
    # PARSE COLUMN CONFIG
    # =========================================================
    def parse_column_config(self, col_config: Dict) -> Dict:
        """
        Normalize column config → unified structure
        """

        if not isinstance(col_config, dict):
            return {"type": "invalid"}

        # TEAM FORMULA
        if col_config.get("type") == "team_formula":
            return {
                "type": "team_formula",
                "team_configs": col_config.get("team_configs", {}),
                "description": col_config.get("description", "")
            }

        # FORMULA
        if "formula" in col_config:
            return {
                "type": "formula",
                "formula": col_config["formula"],
                "description": col_config.get("description", "")
            }

        # SOURCE FILE
        if "source_file" in col_config:
            return {
                "type": "source",
                "file": col_config["source_file"],
                "sheet": col_config.get("source_sheet"),
                "range": col_config.get("source_range"),
                "column_index": col_config.get("column_index"),
                "description": col_config.get("description", "")
            }

        return {"type": "unknown"}

    # =========================================================
    # VALIDATION
    # =========================================================
    def validate_template(self, template_key: str) -> Tuple[bool, List[str]]:
        errors = []

        template = self.get_template_config(template_key)
        if not template:
            return False, [f"Template not found: {template_key}"]

        sheets = template.get("sheets", {})

        for sheet_name, sheet in sheets.items():
            columns = sheet.get("columns", {})

            for col, cfg in columns.items():
                parsed = self.parse_column_config(cfg)

                if parsed["type"] in ("unknown", "invalid"):
                    errors.append(
                        f"{sheet_name}.{col}: invalid config"
                    )

        return len(errors) == 0, errors


# =========================================================
# SINGLETON CONFIG (GLOBAL USAGE)
# =========================================================
_global_config = None


def get_salary_config(config_path: str = None) -> SalarySourceConfig:
    global _global_config

    if _global_config is None:
        _global_config = SalarySourceConfig(config_path)

    return _global_config