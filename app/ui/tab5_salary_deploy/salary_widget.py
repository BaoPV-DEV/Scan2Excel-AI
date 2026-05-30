import os
import json
import openpyxl
from datetime import datetime

from PySide6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QLabel,
    QPushButton, QComboBox, QTextEdit, QMessageBox, QFrame, QProgressBar,
    QGroupBox, QLineEdit, QFileDialog, QScrollArea, QStackedWidget, QSizePolicy
)
from PySide6.QtCore import Qt, Signal
from PySide6.QtGui import QFont

from app.ui.tab5_salary_deploy.salary_thread import SalaryDeployThread


# =========================
# CONFIG
# =========================
SOURCE_FILE_DEFINITIONS = [
    {
        "key": "danh_sach_cbcnv_luong",
        "label": "📋 DS CBCNV bản dùng làm lương",
        "default_sheets": ["Danh sách"],
    },
    {
        "key": "cham_cong",
        "label": "⏱️ Bảng chấm công tháng",
        "default_sheets": ["Công"],
    },
    {
        "key": "chuyen_khoan_thang_truoc",
        "label": "💳 Chuyển khoản tháng trước",
        "default_sheets": ["Bản gốc T{prev_month}"],
    },
    {
        "key": "phu_cap_con_nho",
        "label": "👶 Phụ cấp con nhỏ",
        "default_sheets": ["Con nhỏ", "Nhóm 6", "Thâm niên", "PC phụ nữ"],
    },
]


def _get_user_config_path():
    from app.utils.paths import get_user_run_config_path
    return get_user_run_config_path()


# =========================
# SHEET CACHE (OPTIMIZATION)
# =========================
_SHEET_CACHE = {}


def load_excel_sheets_cached(path: str):
    """Tránh openpyxl load lặp lại cùng file"""
    if path in _SHEET_CACHE:
        return _SHEET_CACHE[path]

    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    sheets = wb.sheetnames
    wb.close()

    _SHEET_CACHE[path] = sheets
    return sheets


# =========================
# ROW WIDGET
# =========================
class SourceFileRow(QFrame):
    config_changed = Signal()

    def __init__(self, definition: dict, parent=None):
        super().__init__(parent)
        self.definition = definition
        self.key = definition["key"]
        self.sheet_combos = {}
        self._loading = False  # chống spam signal
        self._init_ui()

    def _init_ui(self):
        self.setObjectName("SourceRow")

        layout = QVBoxLayout(self)
        layout.setContentsMargins(10, 8, 10, 8)
        layout.setSpacing(6)

        row1 = QHBoxLayout()

        lbl = QLabel(self.definition["label"])
        lbl.setStyleSheet("font-weight: bold; font-size: 12px; min-width: 240px;")
        row1.addWidget(lbl)

        self.txt_path = QLineEdit()
        self.txt_path.setReadOnly(True)
        self.txt_path.setPlaceholderText("Chưa chọn file...")
        row1.addWidget(self.txt_path, 1)

        self.btn_browse = QPushButton("📂 Chọn file")
        self.btn_browse.clicked.connect(self._browse_file)
        row1.addWidget(self.btn_browse)

        self.btn_clear = QPushButton("✖")
        self.btn_clear.clicked.connect(self._clear_selection)
        row1.addWidget(self.btn_clear)

        layout.addLayout(row1)

        # sheet row
        self.sheet_layout = QHBoxLayout()
        for pattern in self.definition["default_sheets"]:
            self.sheet_layout.addWidget(QLabel(f"Sheet [{pattern}]"))

            combo = QComboBox()
            combo.addItem(f"-- Mặc định: {pattern} --")
            self.sheet_layout.addWidget(combo)
            self.sheet_combos[pattern] = combo

        self.sheet_layout.addStretch()
        layout.addLayout(self.sheet_layout)

    # =========================
    # FILE HANDLING OPTIMIZED
    # =========================
    def _browse_file(self):
        file_path, _ = QFileDialog.getOpenFileName(
            self, "Chọn file Excel", "", "Excel (*.xlsx *.xls)"
        )
        if not file_path:
            return

        self.txt_path.setText(file_path)
        self._load_sheets(file_path)
        self.config_changed.emit()

    def _clear_selection(self):
        self.txt_path.clear()

        self._loading = True
        try:
            for pattern, combo in self.sheet_combos.items():
                combo.clear()
                combo.addItem(f"-- Mặc định: {pattern} --")
        finally:
            self._loading = False

        self.config_changed.emit()

    def _load_sheets(self, path: str):
        try:
            sheets = load_excel_sheets_cached(path)

            self._loading = True
            for pattern, combo in self.sheet_combos.items():
                combo.blockSignals(True)

                combo.clear()
                combo.addItem(f"-- Mặc định: {pattern} --")
                combo.addItems(sheets)

                # auto select (fast O(n))
                pattern_clean = pattern.replace("{prev_month}", "").replace("{mm}", "").strip()

                for i, s in enumerate(sheets):
                    if s.lower() == pattern.lower() or (pattern_clean and pattern_clean.lower() in s.lower()):
                        combo.setCurrentIndex(i + 1)
                        break

                combo.blockSignals(False)

            self._loading = False

        except Exception as e:
            QMessageBox.warning(self, "Error", str(e))

    # =========================
    # CONFIG IO OPTIMIZED
    # =========================
    def get_config(self):
        path = self.txt_path.text().strip()
        if not path:
            return {}

        mapping = {}
        for pattern, combo in self.sheet_combos.items():
            val = combo.currentText()
            if val and not val.startswith("--"):
                mapping[pattern] = val

        return {"file_path": path, "sheet_mapping": mapping}

    def set_config(self, cfg: dict):
        path = cfg.get("file_path", "")
        if not path or not os.path.exists(path):
            return

        self.txt_path.setText(path)
        self._load_sheets(path)

        mapping = cfg.get("sheet_mapping", {})
        for pattern, combo in self.sheet_combos.items():
            if pattern in mapping:
                idx = combo.findText(mapping[pattern])
                if idx >= 0:
                    combo.setCurrentIndex(idx)

        self.config_changed.emit()


# =========================
# MAIN WIDGET
# =========================
class SalaryDeployWidget(QWidget):

    PAGE_CONFIG = 0
    PAGE_DEPLOY = 1

    def __init__(self):
        super().__init__()
        self.thread = None
        self.source_rows = {}

        self.init_ui()
        self._load_user_config()
        self._sync_ui()

    # =========================
    # VALIDATION (FAST)
    # =========================
    def _is_complete(self):
        missing = []
        for d in SOURCE_FILE_DEFINITIONS:
            cfg = self.source_rows[d["key"]].get_config()
            if not cfg.get("file_path") or not os.path.exists(cfg["file_path"]):
                missing.append(d["label"])
        return not missing, missing

    # =========================
    # CONFIG CACHE IO
    # =========================
    def _save_user_config(self):
        path = _get_user_config_path()
        os.makedirs(os.path.dirname(path), exist_ok=True)

        data = {k: v.get_config() for k, v in self.source_rows.items() if v.get_config()}

        with open(path, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)

    def _load_user_config(self):
        path = _get_user_config_path()
        if not os.path.exists(path):
            return

        with open(path, "r", encoding="utf-8") as f:
            data = json.load(f)

        for k, cfg in data.items():
            if k in self.source_rows:
                self.source_rows[k].set_config(cfg)

    # =========================
    # UI SYNC (MINIMAL REDRAW)
    # =========================
    def _sync_ui(self):
        ok, missing = self._is_complete()
        self.btn_continue.setVisible(ok)
        self.lbl_status.setVisible(not ok)

        if not ok:
            self.lbl_status.setText(f"Còn thiếu {len(missing)} file")

    # =========================
    # THREAD RUN
    # =========================
    def start_processing(self):
        ok, missing = self._is_complete()
        if not ok:
            QMessageBox.warning(self, "Missing", "\n".join(missing))
            return

        self._save_user_config()

        self.thread = SalaryDeployThread(
            int(self.combo_year.currentText()),
            int(self.combo_month.currentText()),
            {k: v.get_config() for k, v in self.source_rows.items()}
        )
        self.thread.start()