from PySide6.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QPushButton, QLabel, QFrame, QGridLayout, 
    QLineEdit, QTableWidget, QTableWidgetItem, QHeaderView, QAbstractItemView, QMessageBox
)
from PySide6.QtCore import Qt
from PySide6.QtGui import QIntValidator, QColor
from app.ui.tab2_scan.worker_dialog import WorkerEditorDialog


class DataEditorDialog(QDialog):
    """
    Hộp thoại chính để xem và chỉnh sửa toàn bộ dữ liệu JSON bóc tách được từ AI.
    """

    def __init__(self, json_data, parent=None):
        super().__init__(parent)

        self.data = json_data or {}
        self.inputs = {}

        self.setWindowTitle("SỬA DỮ LIỆU SẢN XUẤT (TRỰC TIẾP)")
        self.setWindowFlags(self.windowFlags() | Qt.WindowMaximizeButtonHint | Qt.WindowMinimizeButtonHint)
        self.setWindowState(Qt.WindowMaximized)

        self.setStyleSheet("""
            QDialog { background-color: #F9FAFB; }
            QLabel { font-weight: bold; color: #374151; }
            QLineEdit { padding: 5px; border: 1px solid #D1D5DB; border-radius: 4px; }
            QTableWidget { background-color: white; gridline-color: #E5E7EB; border: 1px solid #E5E7EB; font-size: 13px; }
            QHeaderView::section { background-color: #F3F4F6; padding: 5px; border: 1px solid #E5E7EB; font-weight: bold; }
            QPushButton#SaveBtn { background-color: #10B981; color: white; padding: 8px 20px; font-weight: bold; border-radius: 5px; }
            QPushButton#CancelBtn { background-color: #EF4444; color: white; padding: 8px 20px; font-weight: bold; border-radius: 5px; }
            QPushButton#AddRowBtn { background-color: #3B82F6; color: white; padding: 5px 15px; border-radius: 4px; font-size: 12px; }
            QPushButton#DelRowBtn { background-color: #9CA3AF; color: white; padding: 5px 15px; border-radius: 4px; font-size: 12px; }
        """)

        self.init_ui()

    # ---------------- UI ----------------
    def init_ui(self):
        layout = QVBoxLayout(self)

        # ===== THÔNG TIN CHUNG =====
        info_group = QFrame()
        info_group.setStyleSheet(
            "QFrame { background-color: white; border-radius: 8px; border: 1px solid #E5E7EB; }"
        )

        info_layout = QGridLayout(info_group)
        thong_tin = self.data.get("thong_tin_chung", {})

        fields = [
            ("Mã hàng:", "ma_hang", 0, 0), ("Tổ sản xuất:", "to_san_xuat", 0, 2),
            ("Thời gian:", "thoi_gian", 1, 0), ("Khách hàng:", "khach_hang", 1, 2),
            ("Loại SP:", "loai_san_pham", 2, 0), ("Đơn vị:", "don_vi", 2, 2),
            ("Thưởng mã hàng mới:", "thuong_ma_hang_moi", 3, 0),
            ("Tổng mục tiêu:", "tong_san_luong_muc_tieu", 3, 2)
        ]

        for label_text, key, r, c in fields:
            info_layout.addWidget(QLabel(label_text), r, c)

            edit = QLineEdit(str(thong_tin.get(key, "0" if "thuong" in key else "N/A")))

            if key in ("thuong_ma_hang_moi", "tong_san_luong_muc_tieu"):
                edit.setValidator(QIntValidator(0, 999999999))

            info_layout.addWidget(edit, r, c + 1)
            self.inputs[key] = edit

        layout.addWidget(QLabel("📝 THÔNG TIN CHUNG:"))
        layout.addWidget(info_group)

        # ===== TABLE =====
        layout.addSpacing(10)
        layout.addWidget(QLabel("📊 DANH SÁCH CÔNG ĐOẠN:"))

        self.table = QTableWidget()
        self.table.setColumnCount(5)
        self.table.setHorizontalHeaderLabels(["STT", "Mô Tả", "Định Mức", "Nhân Sự", "Tổng"])
        self.table.setEditTriggers(QAbstractItemView.DoubleClicked | QAbstractItemView.SelectedClicked)

        header = self.table.horizontalHeader()
        header.setSectionResizeMode(0, QHeaderView.ResizeToContents)
        header.setSectionResizeMode(1, QHeaderView.Stretch)
        header.setSectionResizeMode(2, QHeaderView.ResizeToContents)
        header.setSectionResizeMode(3, QHeaderView.Stretch)
        header.setSectionResizeMode(4, QHeaderView.ResizeToContents)

        cong_doan = self.data.get("danh_sach_cong_doan", [])
        self.table.setRowCount(len(cong_doan))

        for i, cd in enumerate(cong_doan):
            self._set_row(i, cd)

        self.table.cellDoubleClicked.connect(self.on_cell_double_clicked)
        layout.addWidget(self.table)

        # ===== BUTTONS =====
        row_btn = QHBoxLayout()

        add_btn = QPushButton("➕ Thêm công đoạn")
        add_btn.clicked.connect(self.add_step)

        del_btn = QPushButton("🗑️ Xóa công đoạn")
        del_btn.clicked.connect(self.remove_step)

        row_btn.addWidget(add_btn)
        row_btn.addWidget(del_btn)
        row_btn.addStretch()

        layout.addLayout(row_btn)

        # ===== SAVE / CANCEL =====
        btn_layout = QHBoxLayout()

        save_btn = QPushButton("💾 LƯU")
        save_btn.setObjectName("SaveBtn")
        save_btn.clicked.connect(self.accept)

        cancel_btn = QPushButton("❌ HỦY")
        cancel_btn.setObjectName("CancelBtn")
        cancel_btn.clicked.connect(self.reject)

        btn_layout.addStretch()
        btn_layout.addWidget(cancel_btn)
        btn_layout.addWidget(save_btn)

        layout.addLayout(btn_layout)

    # ---------------- CORE OPTIMIZED ----------------

    def _set_row(self, row, cd):
        """Tối ưu set row - giảm overhead QTableWidgetItem"""

        stt = str(cd.get("stt", ""))
        mo_ta = str(cd.get("mo_ta", ""))
        dinh_muc = str(cd.get("dinh_muc_t", ""))

        workers = cd.get("thuc_hien", []) or []

        self.table.setItem(row, 0, QTableWidgetItem(stt))
        self.table.setItem(row, 1, QTableWidgetItem(mo_ta))
        self.table.setItem(row, 2, QTableWidgetItem(dinh_muc))

        self._set_worker_cell(row, workers)

    def _set_worker_cell(self, row, workers):
        """Tối ưu hóa Qt UserRole + tránh crash None"""

        if not workers:
            summary = "(Trống)"
            total = 0
            bg = QColor("#FFF59D")
        else:
            summary = " | ".join(
                f"{w.get('nhan_vien')}_{w.get('ma_nhan_vien')} ({w.get('so_luong')})"
                for w in workers
            )
            total = sum(w.get("so_luong", 0) for w in workers)
            bg = QColor("transparent")

        item = QTableWidgetItem(summary)
        item.setData(Qt.UserRole, workers)
        item.setBackground(bg)

        self.table.setItem(row, 3, item)
        self.table.setItem(row, 4, QTableWidgetItem(str(total)))

    # ---------------- EVENTS ----------------

    def on_cell_double_clicked(self, row, col):
        if col != 3:
            return

        item = self.table.item(row, 3)
        if not item:
            return

        workers = item.data(Qt.UserRole) or []

        dialog = WorkerEditorDialog(
            workers,
            self.table.item(row, 1).text() if self.table.item(row, 1) else "",
            self
        )

        if dialog.exec() == QDialog.Accepted:
            self._set_worker_cell(row, dialog.get_workers())

    def add_step(self):
        current_row = self.table.currentRow()
        row = self.table.rowCount() if current_row < 0 else current_row + 1

        self.table.insertRow(row)

        self.table.setItem(row, 0, QTableWidgetItem(str(row + 1)))
        self.table.setItem(row, 1, QTableWidgetItem("Mới"))
        self.table.setItem(row, 2, QTableWidgetItem("0.0"))
        self._set_worker_cell(row, [])

        self._refresh_stt(row + 1)

    def remove_step(self):
        row = self.table.currentRow()
        if row < 0:
            return

        self.table.removeRow(row)
        self._refresh_stt(row)

    def _refresh_stt(self, start=0):
        """Tối ưu STT update - tránh recreate full objects"""
        for r in range(start, self.table.rowCount()):
            item = self.table.item(r, 0)
            if item:
                item.setText(str(r + 1))

    # ---------------- DATA EXPORT ----------------

    def get_updated_data(self):
        try:
            new_info = {
                k: (int(v.text()) if k in ("tong_san_luong_muc_tieu", "thuong_ma_hang_moi") else v.text())
                for k, v in self.inputs.items()
            }

            result = []

            for r in range(self.table.rowCount()):
                workers = self.table.item(r, 3).data(Qt.UserRole) if self.table.item(r, 3) else []

                result.append({
                    "stt": int(self.table.item(r, 0).text() or 0),
                    "mo_ta": self.table.item(r, 1).text() if self.table.item(r, 1) else "",
                    "dinh_muc_t": float(self.table.item(r, 2).text() or 0.0),
                    "thuc_hien": workers or [],
                    "tong_thuc_hien": sum(w.get("so_luong", 0) for w in (workers or []))
                })

            return {
                "thong_tin_chung": new_info,
                "danh_sach_cong_doan": result
            }

        except Exception:
            return None