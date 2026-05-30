from PySide6.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QPushButton, QLabel,
    QTableWidget, QTableWidgetItem, QHeaderView, QDialogButtonBox
)
from PySide6.QtCore import Qt


class WorkerEditorDialog(QDialog):
    """
    Cho phép chỉnh sửa danh sách nhân viên trong một công đoạn.
    (Optimized: safe access + reduced widget overhead + faster save loop)
    """

    def __init__(self, workers, title_info="", parent=None):
        super().__init__(parent)

        self.setWindowTitle("CHI TIẾT NGƯỜI THỰC HIỆN")
        self.resize(550, 450)

        # copy 1 lần duy nhất (tránh reference mutation side-effect)
        self.workers = list(workers) if workers else []
        self.title_info = title_info

        self._init_ui()

    # =========================
    # UI
    # =========================
    def _init_ui(self):
        layout = QVBoxLayout(self)

        info_lbl = QLabel(f"📍 Đang sửa nhân sự cho: {self.title_info}")
        info_lbl.setStyleSheet(
            "font-weight: bold; color: #2563EB; font-size: 14px; padding: 5px;"
        )
        layout.addWidget(info_lbl)

        self.table = QTableWidget()
        self.table.setColumnCount(3)
        self.table.setHorizontalHeaderLabels(["Nhân Viên", "Mã NV", "Số Lượng"])

        header = self.table.horizontalHeader()
        header.setSectionResizeMode(QHeaderView.Stretch)

        layout.addWidget(self.table)

        self._load_data()

        # actions
        btn_layout = QHBoxLayout()

        add_btn = QPushButton("➕ Thêm người")
        add_btn.clicked.connect(self._add_row)

        remove_btn = QPushButton("➖ Xóa người")
        remove_btn.clicked.connect(self._remove_row)

        btn_layout.addWidget(add_btn)
        btn_layout.addWidget(remove_btn)
        btn_layout.addStretch()

        layout.addLayout(btn_layout)

        # ok/cancel
        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        buttons.accepted.connect(self._save_and_accept)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)

    # =========================
    # DATA LOAD
    # =========================
    def _load_data(self):
        workers = self.workers
        self.table.setRowCount(len(workers))

        for i, w in enumerate(workers):
            self.table.setItem(i, 0, QTableWidgetItem(str(w.get("nhan_vien", ""))))
            self.table.setItem(i, 1, QTableWidgetItem(str(w.get("ma_nhan_vien", ""))))
            self.table.setItem(i, 2, QTableWidgetItem(str(w.get("so_luong", 0))))

    # =========================
    # ACTIONS
    # =========================
    def _add_row(self):
        row = self.table.rowCount()
        self.table.insertRow(row)

        self.table.setItem(row, 0, QTableWidgetItem(""))
        self.table.setItem(row, 1, QTableWidgetItem(""))
        self.table.setItem(row, 2, QTableWidgetItem("0"))

    def _remove_row(self):
        row = self.table.currentRow()
        if row >= 0:
            self.table.removeRow(row)

    # =========================
    # SAVE (OPTIMIZED + SAFE)
    # =========================
    def _save_and_accept(self):
        new_workers = []
        row_count = self.table.rowCount()

        for r in range(row_count):
            item_name = self.table.item(r, 0)
            item_ma = self.table.item(r, 1)
            item_sl = self.table.item(r, 2)

            # safe access (tránh None crash)
            name = item_name.text().strip() if item_name else ""
            ma = item_ma.text().strip() if item_ma else ""
            sl_raw = item_sl.text().strip() if item_sl else "0"

            if not (name or ma):
                continue

            try:
                sl = int(sl_raw)
            except ValueError:
                sl = 0

            new_workers.append({
                "nhan_vien": name,
                "ma_nhan_vien": ma,
                "so_luong": sl
            })

        self.workers = new_workers
        self.accept()

    # =========================
    # PUBLIC API
    # =========================
    def get_workers(self):
        return self.workers