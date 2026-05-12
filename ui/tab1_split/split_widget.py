import os
from datetime import datetime
from PySide6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QGridLayout, QLabel, QLineEdit,
    QPushButton, QComboBox, QTextEdit, QFileDialog, QMessageBox, QFrame, QProgressBar, QCheckBox
)
from PySide6.QtCore import Qt
from PySide6.QtGui import QFont
from ui.tab1_split.split_thread import ExcelProcessThread
from utils.paths import get_split_excel_path

class SplitExcelWidget(QWidget):
    """
    Giao diện người dùng cho chức năng Tách File Nhân Viên (Tab 1).
    """
    def __init__(self):
        # Khởi tạo Widget và cấu hình giao diện
        super().__init__()
        self.setObjectName("MainWidget")
        self.setStyleSheet("""
            QWidget#MainWidget { background-color: white; border-radius: 8px; }
            QLabel { font-size: 13px; font-weight: bold; color: #374151; }
            QLineEdit { 
                padding: 8px; border: 1px solid #D1D5DB; border-radius: 5px; background-color: #F9FAFB; color: #111827;
            }
            QComboBox { 
                padding: 6px; border: 1px solid #D1D5DB; border-radius: 5px; background-color: white; color: #111827;
            }
            QPushButton {
                background-color: #3B82F6; color: white; border: none; border-radius: 6px; padding: 8px 15px; font-size: 13px; font-weight: bold;
            }
            QPushButton:hover { background-color: #2563EB; }
            QPushButton:disabled { background-color: #9CA3AF; }
            QPushButton#ProcessBtn { background-color: #10B981; font-size: 15px; padding: 12px; }
            QPushButton#ProcessBtn:hover { background-color: #059669; }
            QProgressBar {
                border: 1px solid #E5E7EB; border-radius: 5px; text-align: center; background-color: #F3F4F6; color: #111827; font-weight: bold; height: 20px;
            }
            QProgressBar::chunk { background-color: #10B981; border-radius: 4px; }
            QTextEdit {
                background-color: #1E1E1E; color: #10B981; font-family: Consolas, monospace; font-size: 13px; border: 1px solid #374151; border-radius: 6px; padding: 10px;
            }
            QFrame#SectionFrame { border: 1px solid #E5E7EB; border-radius: 8px; background-color: #F9FAFB; padding: 10px; }
        """)
        self.init_ui()

    def init_ui(self):
        # Thiết lập các thành phần giao diện
        layout = QVBoxLayout(self)
        layout.setContentsMargins(20, 20, 20, 20)
        layout.setSpacing(15)

        title = QLabel("✂️ TÁCH FILE NHÂN VIÊN THEO TỔ")
        title.setFont(QFont("Arial", 16, QFont.Bold))
        title.setAlignment(Qt.AlignCenter)
        title.setStyleSheet("color: #1F2937; margin-bottom: 10px;")
        layout.addWidget(title)

        # Phần chọn File đầu vào
        file_frame = QFrame()
        file_frame.setObjectName("SectionFrame")
        file_layout = QGridLayout(file_frame)
        file_layout.addWidget(QLabel("📄 Chọn File Danh Sách NV gốc (.xlsx):"), 0, 0)
        self.input_edit = QLineEdit()
        self.input_edit.setReadOnly(True)
        self.input_edit.setPlaceholderText("Bấm nút 'Chọn File' để tìm đường dẫn...")
        file_layout.addWidget(self.input_edit, 0, 1)
        self.btn_browse_input = QPushButton("📂 Chọn File")
        self.btn_browse_input.clicked.connect(self.browse_input)
        file_layout.addWidget(self.btn_browse_input, 0, 2)

        self.template_path = os.path.join(os.getcwd(), "Template", "to_may_file_mau.xlsx")
        layout.addWidget(file_frame)

        # Phần cấu hình Thời gian
        date_frame = QFrame()
        date_frame.setObjectName("SectionFrame")
        date_layout = QHBoxLayout(date_frame)
        now = datetime.now()
        date_layout.addWidget(QLabel("📅 Chọn Thời Gian Báo Cáo:"))
        date_layout.addWidget(QLabel("Tháng:"))
        self.combo_month = QComboBox()
        self.combo_month.addItems([f"{i:02d}" for i in range(1, 13)])
        self.combo_month.setCurrentText(f"{now.month:02d}")
        date_layout.addWidget(self.combo_month)
        date_layout.addWidget(QLabel("Năm:"))
        self.combo_year = QComboBox()
        self.combo_year.addItems([str(y) for y in range(2024, 2031)])
        self.combo_year.setCurrentText(str(now.year))
        date_layout.addWidget(self.combo_year)
        date_layout.addStretch()
        
        self.check_organize = QCheckBox("Tự động phân loại thư mục theo Năm/Tháng")
        self.check_organize.setChecked(True)
        
        layout.addWidget(date_frame)
        layout.addWidget(self.check_organize)

        # Nút thực hiện chính và Progress Bar
        self.btn_run = QPushButton("🚀 THỰC HIỆN TÁCH FILE")
        self.btn_run.setObjectName("ProcessBtn")
        self.btn_run.clicked.connect(self.start_processing)
        layout.addWidget(self.btn_run)
        
        self.progress_bar = QProgressBar()
        layout.addWidget(self.progress_bar)

        self.log_area = QTextEdit()
        self.log_area.setReadOnly(True)
        layout.addWidget(QLabel("📝 Nhật Ký Hoạt Động:"))
        layout.addWidget(self.log_area, stretch=1)

    def browse_input(self):
        # Mở hộp thoại chọn file Excel đầu vào
        path, _ = QFileDialog.getOpenFileName(self, "Chọn File Danh Sách Nhân Viên", "", "Excel files (*.xlsx *.xls)")
        if path: self.input_edit.setText(path)

    def start_processing(self):
        # Kiểm tra thông tin và bắt đầu luồng chạy ngầm
        if not self.input_edit.text():
            QMessageBox.warning(self, "Thiếu Thông Tin", "Vui lòng chọn File Danh Sách NV đầu vào.")
            return

        month = self.combo_month.currentText()
        year = self.combo_year.currentText()
        output_dir = get_split_excel_path(year, month)

        self.btn_run.setEnabled(False)
        self.log_area.clear()
        self.progress_bar.setValue(0)

        self.thread = ExcelProcessThread(
            self.input_edit.text(), self.template_path, output_dir, month, year, False
        )
        self.thread.log_signal.connect(self.log_area.append)
        self.thread.progress_signal.connect(self.progress_bar.setValue)
        self.thread.finished_signal.connect(self.on_processing_finished)
        self.thread.start()

    def on_processing_finished(self, success, message):
        # Xử lý kết quả sau khi luồng chạy xong
        self.btn_run.setEnabled(True)
        if success:
            QMessageBox.information(self, "Thành công", message)
        else:
            QMessageBox.critical(self, "Lỗi", message)
