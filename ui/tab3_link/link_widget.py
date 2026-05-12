import os
from datetime import datetime
from PySide6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QLabel, QLineEdit,
    QPushButton, QComboBox, QTextEdit, QFileDialog, QMessageBox, QFrame, QProgressBar
)
from PySide6.QtCore import Qt
from PySide6.QtGui import QFont
from ui.tab3_link.link_thread import IntegrationThread
from utils.paths import get_json_data_path, get_split_excel_path

class LinkExcelWidget(QWidget):
    """
    Giao diện người dùng cho chức năng Đẩy Dữ Liệu vào Excel (Tab 3).
    """
    def __init__(self):
        # Tự động tạo thư mục the def __init__():
        # Khởi tạo Widget Tab 3
        super().__init__()
        self.setObjectName("LinkWidget")
        self.setStyleSheet("""
            QWidget#LinkWidget { background-color: white; border-radius: 8px; }
            QLabel { font-size: 13px; font-weight: bold; color: #374151; }
            QLineEdit { padding: 8px; border: 1px solid #D1D5DB; border-radius: 5px; }
            QPushButton { background-color: #3B82F6; color: white; padding: 8px 15px; border-radius: 6px; font-weight: bold; }
            QPushButton#ProcessBtn { background-color: #10B981; font-size: 15px; padding: 12px; }
            QProgressBar { border: 1px solid #E5E7EB; border-radius: 5px; text-align: center; height: 20px; }
            QTextEdit { background-color: #1E1E1E; color: #10B981; font-family: Consolas; padding: 10px; }
        """)
        self.init_ui()

    def init_ui(self):
        # Tự động tạo thư mục the def init_ui():
        # Thiết lập UI
        layout = QVBoxLayout(self)
        title = QLabel("🔗 ĐẨY DỮ LIỆU SẢN LƯỢNG VÀO EXCEL")
        title.setFont(QFont("Arial", 16, QFont.Bold))
        title.setAlignment(Qt.AlignCenter)
        layout.addWidget(title)

        # Chọn thời gian
        date_frame = QFrame()
        date_layout = QHBoxLayout(date_frame)
        now = datetime.now()
        date_layout.addWidget(QLabel("📅 Chọn Tháng/Năm:"))
        self.combo_month = QComboBox()
        self.combo_month.addItems([f"{i:02d}" for i in range(1, 13)])
        self.combo_month.setCurrentText(f"{now.month:02d}")
        self.combo_year = QComboBox()
        self.combo_year.addItems([str(y) for y in range(2024, 2031)])
        self.combo_year.setCurrentText(str(now.year))
        date_layout.addWidget(self.combo_month)
        date_layout.addWidget(self.combo_year)
        layout.addWidget(date_frame)

        # Mật khẩu Excel
        pwd_layout = QHBoxLayout()
        pwd_layout.addWidget(QLabel("🔑 Mật khẩu Excel (Nếu có):"))
        self.pwd_edit = QLineEdit()
        self.pwd_edit.setEchoMode(QLineEdit.Password)
        pwd_layout.addWidget(self.pwd_edit)
        layout.addLayout(pwd_layout)

        # Nút thực hiện
        self.btn_run = QPushButton("🚀 BẮT ĐẦU TÍCH HỢP")
        self.btn_run.setObjectName("ProcessBtn")
        self.btn_run.clicked.connect(self.start_processing)
        layout.addWidget(self.btn_run)
        
        self.progress_bar = QProgressBar()
        layout.addWidget(self.progress_bar)
        
        self.log_area = QTextEdit()
        self.log_area.setReadOnly(True)
        layout.addWidget(self.log_area)

    def start_processing(self):
        # Tự động tạo thư mục the def start_processing():
        # Bắt đầu gom JSON và đẩy vào Excel
        month, year = self.combo_month.currentText(), self.combo_year.currentText()
        json_root = os.path.join("D:/Linh_Salary_Tool", "02_ma_hang", year, month)
        excel_dir = get_split_excel_path(year, month)

        if not os.path.exists(json_root):
            QMessageBox.warning(self, "Lỗi", "Không tìm thấy dữ liệu JSON cho tháng này.")
            return

        json_files = []
        for root, _, files in os.walk(json_root):
            for f in files:
                if f.endswith(".json"): json_files.append(os.path.join(root, f))

        self.btn_run.setEnabled(False)
        self.thread = IntegrationThread(json_files, excel_dir, self.pwd_edit.text())
        self.thread.log_signal.connect(self.log_area.append)
        self.thread.progress_signal.connect(self.progress_bar.setValue)
        self.thread.finished_signal.connect(self.on_finished)
        self.thread.start()

    def on_finished(self, success, message):
        # Tự động tạo thư mục the def on_finished():
        # Kết thúc tích hợp
        self.btn_run.setEnabled(True)
        if success: QMessageBox.information(self, "Xong", message)
        else: QMessageBox.critical(self, "Lỗi", message)
