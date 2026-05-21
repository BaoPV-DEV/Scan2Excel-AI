import os
from datetime import datetime
from PySide6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QLabel,
    QPushButton, QComboBox, QTextEdit, QMessageBox, QFrame, QProgressBar
)
from PySide6.QtCore import Qt
from PySide6.QtGui import QFont
from logic.salary_config_parser import get_salary_config
from ui.tab5_salary_deploy.salary_thread import SalaryDeployThread

class SalaryDeployWidget(QWidget):
    """
    Giao diện người dùng cho chức năng Triển khai Công Lương (Tab 5).
    Áp dụng công thức lương từ JSON config - chỉ cần nhập tháng/năm, tất cả khác tự động.
    
    Workflow:
    1. Input tháng/năm
    2. Nhấn "Triển khai" → Tự động:
       - Tìm folder từ config (D:\Linh_Salary_Tool\01_danh_sach_chia_to\YYYY\MM)
       - Tìm file Excel từ folder
       - Áp dụng template từ config
       - Linking với source files tự động
    3. Nếu file source chưa có → để trống + log warning
    """
    def __init__(self):
        super().__init__()
        self.setObjectName("SalaryWidget")
        self.setStyleSheet("""
            QWidget#SalaryWidget { background-color: white; border-radius: 8px; }
            QLabel { font-size: 13px; font-weight: bold; color: #374151; }
            QComboBox { 
                padding: 8px; border: 1px solid #D1D5DB; border-radius: 5px; background-color: white; color: #111827; font-size: 14px;
            }
            QPushButton {
                background-color: #3B82F6; color: white; border: none; border-radius: 6px; padding: 8px 15px; font-size: 13px; font-weight: bold;
            }
            QPushButton:hover { background-color: #2563EB; }
            QPushButton:disabled { background-color: #9CA3AF; }
            QPushButton#ProcessBtn { background-color: #10B981; font-size: 16px; padding: 15px 20px; }
            QPushButton#ProcessBtn:hover { background-color: #059669; }
            QProgressBar {
                border: 1px solid #E5E7EB; border-radius: 5px; text-align: center; background-color: #F3F4F6; color: #111827; font-weight: bold; height: 22px;
            }
            QProgressBar::chunk { background-color: #10B981; border-radius: 4px; }
            QTextEdit {
                background-color: #1E1E1E; color: #10B981; font-family: Consolas, monospace; font-size: 13px; border: 1px solid #374151; border-radius: 6px; padding: 10px;
            }
            QFrame#SectionFrame { border: 1px solid #E5E7EB; border-radius: 8px; background-color: #F9FAFB; padding: 15px; }
        """)
        self.init_ui()

    def init_ui(self):
        layout = QVBoxLayout(self)
        layout.setContentsMargins(20, 20, 20, 20)
        layout.setSpacing(15)

        # Tiêu đề
        title = QLabel("💰 TRIỂN KHAI CÔNG THỨC LƯƠNG")
        title.setFont(QFont("Arial", 18, QFont.Bold))
        title.setAlignment(Qt.AlignCenter)
        title.setStyleSheet("color: #1F2937; margin-bottom: 5px;")
        layout.addWidget(title)

        # Mô tả ngắn gọn
        desc = QLabel("Chỉ cần chọn tháng/năm, hệ thống tự động triển khai công thức từ config")
        desc.setAlignment(Qt.AlignCenter)
        desc.setStyleSheet("color: #666; font-size: 12px; margin-bottom: 10px;")
        layout.addWidget(desc)

        # Phần nhập Tháng/Năm
        date_frame = QFrame()
        date_frame.setObjectName("SectionFrame")
        date_layout = QHBoxLayout(date_frame)
        
        now = datetime.now()
        
        date_layout.addWidget(QLabel("📅 Tháng:"))
        self.combo_month = QComboBox()
        self.combo_month.addItems([f"{i:02d}" for i in range(1, 13)])
        self.combo_month.setCurrentText(f"{now.month:02d}")
        self.combo_month.setMinimumWidth(80)
        date_layout.addWidget(self.combo_month)
        
        date_layout.addWidget(QLabel("Năm:"))
        self.combo_year = QComboBox()
        self.combo_year.addItems([str(y) for y in range(2024, 2031)])
        self.combo_year.setCurrentText(str(now.year))
        self.combo_year.setMinimumWidth(100)
        date_layout.addWidget(self.combo_year)
        
        date_layout.addStretch()
        layout.addWidget(date_frame)

        # Nút thực hiện chính
        self.btn_run = QPushButton("🚀 TRIỂN KHAI NGAY")
        self.btn_run.setObjectName("ProcessBtn")
        self.btn_run.clicked.connect(self.start_processing)
        self.btn_run.setCursor(Qt.PointingHandCursor)
        layout.addWidget(self.btn_run)
        
        self.progress_bar = QProgressBar()
        self.progress_bar.setVisible(False)
        layout.addWidget(self.progress_bar)
        
        # Nhật ký hoạt động
        log_label = QLabel("📝 Nhật Ký Hoạt Động:")
        log_label.setStyleSheet("color: #374151; font-weight: bold; font-size: 13px; margin-top: 10px;")
        layout.addWidget(log_label)
        
        self.log_area = QTextEdit()
        self.log_area.setReadOnly(True)
        layout.addWidget(self.log_area, stretch=1)

    def start_processing(self):
        month = self.combo_month.currentText()
        year = self.combo_year.currentText()

        self.btn_run.setEnabled(False)
        self.log_area.clear()
        self.progress_bar.setValue(0)
        self.progress_bar.setVisible(True)

        # Load config để lấy folder và templates
        try:
            config = get_salary_config()
            base_path = config.config.get("base_path", "D:\\Linh_Salary_Tool\\01_danh_sach_chia_to")
            templates = list(config.templates.keys())
            
            if not templates:
                QMessageBox.critical(self, "Lỗi", "Không tìm thấy template nào trong config")
                self.btn_run.setEnabled(True)
                self.progress_bar.setVisible(False)
                return
            
            # Folder output = base_path + YYYY\MM
            excel_files_dir = os.path.join(base_path, year, month)
            
            # Log info
            self.log_area.append(f"🔧 Cấu hình:")
            self.log_area.append(f"  📁 Folder: {excel_files_dir}")
            self.log_area.append(f"  📋 Templates: {', '.join(templates)}")
            self.log_area.append(f"  📅 Tháng/Năm: {month}/{year}")
            self.log_area.append(f"\n{'='*60}\n")
            
            # Triển khai từng template
            all_success = True
            for template_key in templates:
                self.log_area.append(f"⏳ Triển khai: {template_key}...")
                
                self.thread = SalaryDeployThread(
                    excel_files_dir,
                    template_key,
                    month,
                    year
                )
                self.thread.log_signal.connect(self.log_area.append)
                self.thread.progress_signal.connect(self.progress_bar.setValue)
                self.thread.finished_signal.connect(
                    lambda success, msg, tmpl=template_key: self.on_template_finished(success, msg, tmpl)
                )
                self.thread.start()
                # Wait for thread to finish
                self.thread.wait()
                
        except Exception as e:
            self.log_area.append(f"❌ Lỗi: {str(e)}")
            import traceback
            self.log_area.append(traceback.format_exc())
            self.btn_run.setEnabled(True)
            self.progress_bar.setVisible(False)

    def on_template_finished(self, success, message, template_key):
        if success:
            self.log_area.append(f"✅ {template_key}: {message}\n")
        else:
            self.log_area.append(f"⚠️ {template_key}: {message}\n")
        
        self.btn_run.setEnabled(True)
        self.progress_bar.setVisible(False)

