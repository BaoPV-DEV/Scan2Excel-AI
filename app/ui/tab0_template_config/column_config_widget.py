import os
import json

from PySide6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QLabel, QLineEdit,
    QPushButton, QTableWidget, QTableWidgetItem, QHeaderView,
    QMessageBox, QProgressBar, QTextEdit, QFrame, QSizePolicy
)
from PySide6.QtCore import Qt, Slot
from PySide6.QtGui import QFont, QColor

from app.ui.tab0_template_config.column_config_thread import TemplateUpdateThread
from app.logic.template_updater import get_column_letter


class ColumnConfigWidget(QWidget):

    DEFAULT_COLUMNS = [
        "Tổ 1", "Tổ 2", "Tổ 3", "Tổ 4", "Tổ 5", "Tổ 6",
        "Tổ 7", "Tổ 8", "Tổ 9", "Tổ 10", "Tổ 11", "Tổ 12",
        "Là TP", "Kiểm hoá", "Cơ động", "Cắt", "Hoàn thiện", "Tái chế"
    ]

    MAX_COLUMNS = 24
    START_COL_INDEX = 6  # F

    def __init__(self):
        super().__init__()
        self.setObjectName("ColumnConfigWidget")
        self.setAttribute(Qt.WA_StyledBackground, True)

        from app.utils.paths import get_template_columns_path
        self.config_path = get_template_columns_path()

        self.columns = []
        self._loading = False

        # cache mapping để giảm tính toán lại
        self._mapping_cache = {}

        self.load_columns_from_json()
        self._init_ui()

    # =========================
    # UI (GIỮ NGUYÊN)
    # =========================
    def _init_ui(self):
        layout = QVBoxLayout(self)
        layout.setContentsMargins(15, 15, 15, 15)
        layout.setSpacing(10)

        title = QLabel("🛠️ CẤU HÌNH CỘT TIÊU ĐỀ TEMPLATE")
        title.setFont(QFont("Arial", 16, QFont.Bold))
        title.setAlignment(Qt.AlignCenter)
        layout.addWidget(title)

        desc = QLabel(
            "Mỗi tháng công ty có thể thay đổi danh sách tổ hoặc bộ phận."
        )
        desc.setAlignment(Qt.AlignCenter)
        layout.addWidget(desc)

        main = QHBoxLayout()

        # TABLE
        self.table = QTableWidget()
        self.table.setColumnCount(3)
        self.table.setHorizontalHeaderLabels(["STT", "Cột Excel", "Tên Tiêu Đề"])

        header = self.table.horizontalHeader()
        header.setSectionResizeMode(0, QHeaderView.ResizeToContents)
        header.setSectionResizeMode(1, QHeaderView.ResizeToContents)
        header.setSectionResizeMode(2, QHeaderView.Stretch)

        self.table.setSelectionBehavior(QTableWidget.SelectRows)
        self.table.setSelectionMode(QTableWidget.SingleSelection)
        self.table.itemSelectionChanged.connect(self.on_selection_changed)

        main.addWidget(self.table, 3)

        # PANEL
        panel = QFrame()
        panel_layout = QVBoxLayout(panel)

        self.txt_title = QLineEdit()
        self.txt_title.returnPressed.connect(self.add_column)

        panel_layout.addWidget(QLabel("Tiêu đề"))
        panel_layout.addWidget(self.txt_title)

        self.btn_add = QPushButton("➕ Thêm")
        self.btn_add.clicked.connect(self.add_column)

        self.btn_edit = QPushButton("Sửa")
        self.btn_edit.clicked.connect(self.edit_column)
        self.btn_edit.setEnabled(False)

        self.btn_delete = QPushButton("Xóa")
        self.btn_delete.clicked.connect(self.delete_column)
        self.btn_delete.setEnabled(False)

        self.btn_up = QPushButton("↑")
        self.btn_up.clicked.connect(self.move_up)
        self.btn_up.setEnabled(False)

        self.btn_down = QPushButton("↓")
        self.btn_down.clicked.connect(self.move_down)
        self.btn_down.setEnabled(False)

        self.btn_reset = QPushButton("Reset")
        self.btn_reset.clicked.connect(self.reset_to_defaults)

        for b in (
            self.btn_add, self.btn_edit, self.btn_delete,
            self.btn_up, self.btn_down, self.btn_reset
        ):
            panel_layout.addWidget(b)

        panel_layout.addStretch()
        main.addWidget(panel, 1)

        layout.addLayout(main)

        self.btn_run = QPushButton("CẬP NHẬT TEMPLATE")
        self.btn_run.clicked.connect(self.start_updating_templates)
        layout.addWidget(self.btn_run)

        self.progress_bar = QProgressBar()
        layout.addWidget(self.progress_bar)

        self.log_area = QTextEdit()
        self.log_area.setReadOnly(True)
        layout.addWidget(self.log_area)

        # UI STYLE (GIỮ NGUYÊN BẢN CŨ)
        self.setStyleSheet("""
            QWidget#ColumnConfigWidget { background-color: white; border-radius: 8px; }
            QLabel { font-size: 13px; font-weight: bold; color: #374151; }
            QLineEdit { padding: 8px; border: 1px solid #D1D5DB; border-radius: 5px; font-size: 13px; }

            QPushButton { background-color: #3B82F6; color: white; padding: 8px 15px;
                          border-radius: 6px; font-weight: bold; font-size: 13px; border: none; }

            QPushButton:hover { background-color: #2563EB; }
            QPushButton:pressed { background-color: #1D4ED8; }

            QPushButton#DeleteBtn { background-color: #EF4444; }
            QPushButton#DeleteBtn:hover { background-color: #DC2626; }

            QPushButton#ResetBtn { background-color: #F59E0B; }
            QPushButton#ResetBtn:hover { background-color: #D97706; }

            QPushButton#ProcessBtn { background-color: #10B981; font-size: 15px;
                                      padding: 12px; border-radius: 8px; }

            QPushButton#ProcessBtn:hover { background-color: #059669; }

            QProgressBar { border: 1px solid #E5E7EB; border-radius: 5px;
                           text-align: center; height: 20px; font-weight: bold; }

            QTextEdit { background-color: #1E1E1E; color: #10B981;
                        font-family: Consolas; padding: 10px; font-size: 12px; }

            QTableWidget { border: 1px solid #E5E7EB; border-radius: 5px;
                           background-color: white; gridline-color: #F3F4F6; }

            QTableWidget::item { padding: 5px; }

            QHeaderView::section { background-color: #F3F4F6; color: #374151;
                                   font-weight: bold; border: 1px solid #E5E7EB; }
        """)

        self.refresh_table()

    # =========================
    # JSON
    # =========================
    def load_columns_from_json(self):
        if not os.path.exists(self.config_path):
            self.columns = self.DEFAULT_COLUMNS.copy()
            self.save_columns_to_json()
            return

        try:
            with open(self.config_path, "r", encoding="utf-8") as f:
                data = json.load(f)
            self.columns = data.get("columns", self.DEFAULT_COLUMNS.copy())
        except Exception:
            self.columns = self.DEFAULT_COLUMNS.copy()

    def save_columns_to_json(self):
        os.makedirs(os.path.dirname(self.config_path), exist_ok=True)

        self._mapping_cache = {
            name: get_column_letter(self.START_COL_INDEX + i)
            for i, name in enumerate(self.columns)
        }

        try:
            with open(self.config_path, "w", encoding="utf-8") as f:
                json.dump(
                    {"columns": self.columns, "mapping": self._mapping_cache},
                    f,
                    ensure_ascii=False,
                    indent=2
                )
        except Exception as e:
            QMessageBox.critical(self, "Error", str(e))

    # =========================
    # TABLE OPTIMIZED RENDER
    # =========================
    def refresh_table(self):
        self._loading = True
        self.table.setUpdatesEnabled(False)
        self.table.blockSignals(True)

        self.table.setRowCount(0)

        for i, name in enumerate(self.columns):
            r = self.table.rowCount()
            self.table.insertRow(r)

            self._set_item(r, 0, str(r + 1), center=True)
            self._set_item(r, 1, get_column_letter(self.START_COL_INDEX + i), color="#2563EB")
            self._set_item(r, 2, name)

        self.table.blockSignals(False)
        self.table.setUpdatesEnabled(True)

        self._loading = False
        self._reset_buttons()

    def _set_item(self, row, col, value, center=False, color=None):
        item = QTableWidgetItem(value)
        item.setFlags(item.flags() & ~Qt.ItemIsEditable)

        if center:
            item.setTextAlignment(Qt.AlignCenter)

        if color:
            item.setForeground(QColor(color))

        self.table.setItem(row, col, item)

    def _reset_buttons(self):
        self.txt_title.clear()
        self.btn_edit.setEnabled(False)
        self.btn_delete.setEnabled(False)
        self.btn_up.setEnabled(False)
        self.btn_down.setEnabled(False)

    # =========================
    # EVENTS
    # =========================
    def on_selection_changed(self):
        if self._loading:
            return

        rows = self.table.selectionModel().selectedRows()
        if not rows:
            self._reset_buttons()
            return

        r = rows[0].row()
        self.txt_title.setText(self.columns[r])

        self.btn_edit.setEnabled(True)
        self.btn_delete.setEnabled(True)
        self.btn_up.setEnabled(r > 0)
        self.btn_down.setEnabled(r < len(self.columns) - 1)

    # =========================
    # ACTIONS
    # =========================
    @Slot()
    def add_column(self):
        text = self.txt_title.text().strip()
        if not text or len(self.columns) >= self.MAX_COLUMNS or text in self.columns:
            return

        self.columns.append(text)
        self.save_columns_to_json()
        self.refresh_table()

    @Slot()
    def edit_column(self):
        rows = self.table.selectionModel().selectedRows()
        if not rows:
            return

        r = rows[0].row()
        new_text = self.txt_title.text().strip()
        if not new_text:
            return

        self.columns[r] = new_text
        self.save_columns_to_json()
        self.refresh_table()

    @Slot()
    def delete_column(self):
        rows = self.table.selectionModel().selectedRows()
        if not rows:
            return

        self.columns.pop(rows[0].row())
        self.save_columns_to_json()
        self.refresh_table()

    @Slot()
    def move_up(self):
        r = self.table.selectionModel().selectedRows()[0].row()
        if r == 0:
            return

        self.columns[r], self.columns[r - 1] = self.columns[r - 1], self.columns[r]
        self.save_columns_to_json()
        self.refresh_table()
        self.table.selectRow(r - 1)

    @Slot()
    def move_down(self):
        r = self.table.selectionModel().selectedRows()[0].row()
        if r >= len(self.columns) - 1:
            return

        self.columns[r], self.columns[r + 1] = self.columns[r + 1], self.columns[r]
        self.save_columns_to_json()
        self.refresh_table()
        self.table.selectRow(r + 1)

    @Slot()
    def reset_to_defaults(self):
        self.columns = self.DEFAULT_COLUMNS.copy()
        self.save_columns_to_json()
        self.refresh_table()

    @Slot()
    def start_updating_templates(self):
        if not self.columns:
            return

        self.thread = TemplateUpdateThread(self.columns)
        self.thread.log_signal.connect(self.log_area.append)
        self.thread.progress_signal.connect(self.progress_bar.setValue)
        self.thread.start()