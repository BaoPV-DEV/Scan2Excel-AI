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
        from app.utils.paths import get_salary_sources_path, get_base_path
        if config_path is None:
            config_path = get_salary_sources_path()
        
        self.config_path = config_path
        self.config = self._load_config()
        self.base_path = self.config.get("base_path", get_base_path())
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
        Returns dict with 'type' (formula, team_formula, or source), and corresponding data.
        """
        if "type" in col_config and col_config["type"] == "team_formula":
            return {
                "type": "team_formula",
                "team_configs": col_config.get("team_configs", {}),
                "description": col_config.get("description", "")
            }
        elif "formula" in col_config:
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
