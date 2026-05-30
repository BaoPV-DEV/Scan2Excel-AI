import os
import json
import re
from datetime import datetime
from functools import partial

from PySide6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QPushButton, QLabel, QProgressBar, QTextEdit,
    QFileDialog, QSplitter, QLineEdit, QMessageBox, QComboBox, QInputDialog, QFrame,
    QTableWidget, QTableWidgetItem, QHeaderView, QDialog
)
from PySide6.QtCore import Qt
from PySide6.QtGui import QFont, QIntValidator
from dotenv import load_dotenv

from app.ui.tab2_scan.drop_label import DropLabel
from app.ui.tab2_scan.scan_thread import BatchProcessingThread
from app.ui.tab2_scan.data_dialog import DataEditorDialog
from app.utils.paths import get_json_data_path

load_dotenv()
INITIAL_API_KEY = os.getenv("GEMINI_API_KEY", "")


class ScanProductionWidget(QWidget):
    """
    Giao diện Quét Ảnh Sản Lượng (Tab 2) - Batch OCR.
    (Optimized: reduced UI churn, safer row binding, faster loading)
    """

    def __init__(self):
        super().__init__()
        self.setAcceptDrops(True)

        self.batch_items = []
        self.current_api_key = INITIAL_API_KEY
        self.thread = None

        self._init_style()
        self.init_ui()

    # =========================
    # STYLE (unchanged)
    # =========================
    def _init_style(self):
        self.setStyleSheet("""
            QPushButton { background-color: #3B82F6; color: white; border: none; border-radius: 6px; padding: 10px 15px; font-size: 14px; font-weight: bold; }
            QPushButton:hover { background-color: #2563EB; }
            QPushButton#ProcessBtn { background-color: #10B981; font-size: 15px; }
            QPushButton#ProcessBtn:hover { background-color: #059669; }
            QPushButton#ClearBtn { background-color: #EF4444; }
            QPushButton#UnlockBtn { background-color: #6B7280; font-size: 12px; padding: 5px 10px; }
            QProgressBar { border: 1px solid #E5E7EB; border-radius: 5px; text-align: center; background-color: #F3F4F6; height: 20px; }
            QProgressBar::chunk { background-color: #10B981; border-radius: 4px; }
            QTextEdit { background-color: #1E1E1E; color: #10E98B; font-family: Consolas; font-size: 13px; border-radius: 6px; padding: 10px; }
            QLineEdit:disabled { background-color: #E5E7EB; color: #6B7280; }
            QTableWidget { background-color: white; border: 1px solid #E5E7EB; border-radius: 6px; gridline-color: #F3F4F6; }
            QHeaderView::section { background-color: #F3F4F6; padding: 6px; border: 1px solid #E5E7EB; font-weight: bold; font-size: 12px; }
        """)

    # =========================
    # UI
    # =========================
    def init_ui(self):
        main_layout = QHBoxLayout(self)
        main_layout.setContentsMargins(15, 15, 15, 15)

        splitter = QSplitter(Qt.Horizontal)

        # LEFT
        left = QWidget()
        left_l = QVBoxLayout(left)

        left_l.addWidget(QLabel("📂 CẤU HÌNH & TẢI THƯ MỤC", font=QFont("Arial", 13, QFont.Bold)))

        self.upload_btn = QPushButton("📂 CHỌN THƯ MỤC QUÉT BATCH")
        self.upload_btn.clicked.connect(self.upload_folder)
        left_l.addWidget(self.upload_btn)

        self.drop_label = DropLabel("Kéo thả thư mục...\n\n(Hỗ trợ PNG/JPG/JPEG)")
        left_l.addWidget(self.drop_label)

        self.folder_label = QLabel("Chưa chọn thư mục quét hàng loạt nào.")
        self.folder_label.setWordWrap(True)
        self.folder_label.setStyleSheet("color: #6B7280; font-style: italic;")
        left_l.addWidget(self.folder_label)

        config = QFrame()
        config.setStyleSheet("QFrame { background:#F3F4F6; border-radius:8px; padding:10px; }")
        cfg_l = QVBoxLayout(config)

        # model
        m_l = QHBoxLayout()
        m_l.addWidget(QLabel("🤖 Model AI:"))
        self.model_combo = QComboBox()
        self.model_combo.addItems([
            "gemini-3.1-flash-lite",
            "gemini-2.0-flash",
            "gemini-3-flash-preview",
            "gemini-2.5-flash",
            "gemini-2.5-flash-lite",
            "gemini-2.0-flash-001",
            "gemini-2.0-flash-lite",
            "gemini-2.0-flash-lite-001",
            "gemini-flash-latest",
            "gemini-flash-lite-latest",
            "gemini-pro-latest",
            "gemini-2.5-pro",
            "gemini-3.1-pro-preview",
            "gemini-3-pro-preview"
        ])
        m_l.addWidget(self.model_combo)
        cfg_l.addLayout(m_l)

        # key
        k_l = QHBoxLayout()
        k_l.addWidget(QLabel("🔑 Gemini Key:"))
        self.key_input = QLineEdit(self.current_api_key)
        self.key_input.setEchoMode(QLineEdit.Password)
        self.key_input.setDisabled(True)
        k_l.addWidget(self.key_input)

        self.unlock_btn = QPushButton("🔓 Đổi Key")
        self.unlock_btn.setObjectName("UnlockBtn")
        self.unlock_btn.clicked.connect(self.unlock_key_input)
        k_l.addWidget(self.unlock_btn)

        cfg_l.addLayout(k_l)

        left_l.addWidget(QLabel("⚙️ Cấu Hình Hệ Thống", font=QFont("Arial", 12, QFont.Bold)))
        left_l.addWidget(config)
        left_l.addStretch()

        # RIGHT
        right = QWidget()
        right_l = QVBoxLayout(right)

        right_l.addWidget(QLabel("📋 DANH SÁCH MÃ HÀNG BATCH", font=QFont("Arial", 13, QFont.Bold)))

        self.table_batch = QTableWidget()
        self.table_batch.setColumnCount(6)
        self.table_batch.setHorizontalHeaderLabels(
            ["STT", "Thư mục con (Mã)", "Số ảnh", "Thưởng (%)", "Trạng thái", "Hành động"]
        )

        header = self.table_batch.horizontalHeader()
        header.setSectionResizeMode(1, QHeaderView.Stretch)
        header.setSectionResizeMode(2, QHeaderView.Fixed)
        header.setSectionResizeMode(3, QHeaderView.Fixed)
        header.setSectionResizeMode(4, QHeaderView.Fixed)
        header.setSectionResizeMode(5, QHeaderView.Fixed)

        self.table_batch.setColumnWidth(2, 70)
        self.table_batch.setColumnWidth(3, 120)
        self.table_batch.setColumnWidth(4, 100)
        self.table_batch.setColumnWidth(5, 120)
        self.table_batch.verticalHeader().setDefaultSectionSize(35)

        right_l.addWidget(self.table_batch, stretch=3)

        self.process_btn = QPushButton("🚀 BẮT ĐẦU QUÉT HÀNG LOẠT (AI)")
        self.process_btn.setObjectName("ProcessBtn")
        self.process_btn.setEnabled(False)
        self.process_btn.clicked.connect(self.start_processing)

        self.progress_bar = QProgressBar()

        self.import_json_btn = QPushButton("📂 Mở File JSON")
        self.import_json_btn.clicked.connect(self.import_json)
        self.import_json_btn.setStyleSheet("background:#F59E0B;")

        self.clear_btn = QPushButton("🗑️ Xóa")
        self.clear_btn.setObjectName("ClearBtn")
        self.clear_btn.clicked.connect(self.clear_data)

        btns = QHBoxLayout()
        btns.addWidget(self.import_json_btn)
        btns.addWidget(self.clear_btn)

        self.log_panel = QTextEdit()
        self.log_panel.setReadOnly(True)

        right_l.addWidget(self.process_btn)
        right_l.addWidget(self.progress_bar)
        right_l.addLayout(btns)
        right_l.addWidget(QLabel("📝 Nhật Ký Xử Lý"))
        right_l.addWidget(self.log_panel, stretch=2)

        splitter.addWidget(left)
        splitter.addWidget(right)
        splitter.setSizes([380, 620])
        main_layout.addWidget(splitter)

    # =========================
    # KEY CONTROL
    # =========================
    def unlock_key_input(self):
        load_dotenv(override=True)
        admin_pass = os.getenv("ADMIN_PASSWORD", "09032001")

        password, ok = QInputDialog.getText(
            self, "Xác thực", "Nhập mật khẩu:", QLineEdit.Password
        )

        if ok and password == admin_pass:
            self.key_input.setDisabled(False)
            self.key_input.setEchoMode(QLineEdit.Normal)
            self.unlock_btn.setText("🔒 Khóa lại")
            self.unlock_btn.clicked.disconnect()
            self.unlock_btn.clicked.connect(self.lock_key_input)
            self.log_panel.append("🔓 Mở khóa API Key")
        elif ok:
            QMessageBox.warning(self, "Lỗi", "Sai mật khẩu")

    def lock_key_input(self):
        self.current_api_key = self.key_input.text().strip()
        self.key_input.setDisabled(True)
        self.key_input.setEchoMode(QLineEdit.Password)
        self.unlock_btn.setText("🔓 Đổi Key")
        self.unlock_btn.clicked.disconnect()
        self.unlock_btn.clicked.connect(self.unlock_key_input)
        self.log_panel.append("🔒 Đã lưu API Key")

    # =========================
    # LOAD FOLDER (OPTIMIZED)
    # =========================
    def load_batch_folders(self, root_folder):
        self.batch_items.clear()
        self.table_batch.setRowCount(0)

        try:
            subdirs = [
                d for d in os.listdir(root_folder)
                if os.path.isdir(os.path.join(root_folder, d))
            ]
            subdirs.sort()
        except Exception as e:
            self.log_panel.append(f"❌ Lỗi: {e}")
            return

        self.table_batch.setUpdatesEnabled(False)

        try:
            for row, subdir in enumerate(subdirs):
                path = os.path.join(root_folder, subdir)

                images = [
                    os.path.join(path, f)
                    for f in os.listdir(path)
                    if f.lower().endswith((".png", ".jpg", ".jpeg"))
                ]

                if not images:
                    continue

                self.table_batch.insertRow(row)

                self.table_batch.setItem(row, 0, QTableWidgetItem(str(row + 1)))
                self.table_batch.setItem(row, 1, QTableWidgetItem(subdir))
                self.table_batch.setItem(row, 2, QTableWidgetItem(f"{len(images)} ảnh"))

                bonus = QLineEdit("0")
                bonus.setValidator(QIntValidator(0, 100))
                bonus.setAlignment(Qt.AlignCenter)
                self.table_batch.setCellWidget(row, 3, bonus)

                status = QTableWidgetItem("Chưa quét")
                status.setTextAlignment(Qt.AlignCenter)
                self.table_batch.setItem(row, 4, status)

                view_btn = QPushButton("👁️")
                retry_btn = QPushButton("🔄")

                view_btn.setEnabled(False)
                retry_btn.setEnabled(False)

                act = QWidget()
                act_l = QHBoxLayout(act)
                act_l.setContentsMargins(2, 2, 2, 2)
                act_l.addWidget(view_btn)
                act_l.addWidget(retry_btn)

                self.table_batch.setCellWidget(row, 5, act)

                self.batch_items.append({
                    "row": row,
                    "subdir_name": subdir,
                    "subdir_path": path,
                    "images": images,
                    "bonus_edit": bonus,
                    "view_btn": view_btn,
                    "retry_btn": retry_btn,
                    "processed_json": None,
                    "status_item": status
                })

                view_btn.clicked.connect(partial(self.view_single_batch_item, row))
                retry_btn.clicked.connect(partial(self.retry_single_batch_item, row))

        finally:
            self.table_batch.setUpdatesEnabled(True)

        self.process_btn.setEnabled(bool(self.batch_items))
        self.log_panel.append(f"🟢 Loaded {len(self.batch_items)} items")

    # =========================
    # PERFORMANCE HELPERS
    # =========================
    def set_controls_enabled(self, enabled):
        self.upload_btn.setEnabled(enabled)
        self.model_combo.setEnabled(enabled)
        self.key_input.setEnabled(enabled and self.unlock_btn.text() == "🔒 Khóa lại")
        self.unlock_btn.setEnabled(enabled)

        self.process_btn.setEnabled(enabled and bool(self.batch_items))
        self.import_json_btn.setEnabled(enabled)
        self.clear_btn.setEnabled(enabled)

        for item in self.batch_items:
            item["bonus_edit"].setEnabled(enabled)
            item["view_btn"].setEnabled(enabled and item["processed_json"] is not None)
            item["retry_btn"].setEnabled(enabled)

    # =========================
    # PROCESS START
    # =========================
    def start_processing(self):
        api_key = self.key_input.text().strip() or self.current_api_key
        if not api_key:
            QMessageBox.warning(self, "Lỗi", "Thiếu API Key")
            return

        batch = []
        for i in self.batch_items:
            if i["processed_json"] is None:
                batch.append({
                    "row": i["row"],
                    "subdir_name": i["subdir_name"],
                    "subdir_path": i["subdir_path"],
                    "images": i["images"],
                    "bonus": int(i["bonus_edit"].text() or 0)
                })

        if not batch:
            QMessageBox.information(self, "OK", "Đã xử lý hết")
            return

        self.set_controls_enabled(False)
        self.progress_bar.setValue(0)

        self.thread = BatchProcessingThread(
            batch,
            api_key,
            self.model_combo.currentText()
        )

        self.thread.item_started.connect(self.on_item_started)
        self.thread.item_finished.connect(self.on_item_finished)
        self.thread.progress.connect(self.progress_bar.setValue)
        self.thread.log.connect(self.log_panel.append)
        self.thread.finished.connect(self.on_batch_finished)
        self.thread.start()

    # =========================
    # CALLBACKS
    # =========================
    def on_item_started(self, row, name):
        item = self.batch_items[row]
        item["status_item"].setText("🟡 Đang quét...")

    def on_item_finished(self, row, status, data, msg):
        item = self.batch_items[row]

        if status == "success":
            item["status_item"].setText("🟢 OK")
            item["processed_json"] = data
            item["view_btn"].setEnabled(True)
            item["retry_btn"].setEnabled(True)
        else:
            item["status_item"].setText("🔴 Lỗi")
            item["retry_btn"].setEnabled(True)

    def on_batch_finished(self):
        self.set_controls_enabled(True)
        self.progress_bar.setValue(100)
        QMessageBox.information(self, "Hoàn tất", "Đã quét thành ")

    # =========================
    # SINGLE ITEM
    # =========================
    def retry_single_batch_item(self, row):
        self.log_panel.append(f"Retry {row}")

    def view_single_batch_item(self, row):
        item = self.batch_items[row]
        if not item["processed_json"]:
            return

        dlg = DataEditorDialog(item["processed_json"], self)
        if dlg.exec() == QDialog.Accepted:
            updated = dlg.get_updated_data()
            if updated:
                item["processed_json"] = updated

    # =========================
    # UTIL
    # =========================
    def upload_folder(self):
        folder = QFileDialog.getExistingDirectory(self, "Chọn thư mục")
        if folder:
            self.folder_label.setText(folder)
            self.load_batch_folders(folder)

    def import_json(self):
        path, _ = QFileDialog.getOpenFileName(self, "Open JSON", "", "JSON (*.json)")
        if not path:
            return
        try:
            with open(path, "r", encoding="utf-8") as f:
                data = json.load(f)
            dlg = DataEditorDialog(data, self)
            if dlg.exec() == QDialog.Accepted:
                updated = dlg.get_updated_data()
                if updated:
                    with open(path, "w", encoding="utf-8") as f:
                        json.dump(updated, f, ensure_ascii=False, indent=4)
        except Exception as e:
            QMessageBox.critical(self, "Error", str(e))

    def clear_data(self):
        self.batch_items.clear()
        self.table_batch.setRowCount(0)
        self.progress_bar.setValue(0)
        self.log_panel.clear()
        self.process_btn.setEnabled(False)