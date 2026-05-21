import os
from datetime import datetime
from PySide6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QLabel, QLineEdit,
    QPushButton, QComboBox, QTextEdit, QFileDialog, QMessageBox, QFrame, QProgressBar, QRadioButton
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

        # Chế độ chạy
        self.mode_frame = QFrame()
        mode_layout = QHBoxLayout(self.mode_frame)
        self.radio_all = QRadioButton("Chạy tất cả các tổ")
        self.radio_single = QRadioButton("Chạy 1 tổ")
        self.radio_all.setChecked(True)
        mode_layout.addWidget(QLabel("⚙️ Chế độ:"))
        mode_layout.addWidget(self.radio_all)
        mode_layout.addWidget(self.radio_single)
        layout.addWidget(self.mode_frame)

        # Chọn thời gian
        self.date_frame = QFrame()
        date_layout = QHBoxLayout(self.date_frame)
        now = datetime.now()
        date_layout.addWidget(QLabel("📅 Chọn Tháng/Năm (Gốc quét):"))
        self.combo_month = QComboBox()
        self.combo_month.addItems([f"{i:02d}" for i in range(1, 13)])
        self.combo_month.setCurrentText(f"{now.month:02d}")
        self.combo_year = QComboBox()
        self.combo_year.addItems([str(y) for y in range(2024, 2031)])
        self.combo_year.setCurrentText(str(now.year))
        date_layout.addWidget(self.combo_month)
        date_layout.addWidget(self.combo_year)
        layout.addWidget(self.date_frame)

        # Chọn Folder 1 tổ
        self.folder_frame = QFrame()
        folder_layout = QHBoxLayout(self.folder_frame)
        self.folder_path = QLineEdit()
        self.folder_path.setPlaceholderText("Chọn folder chứa file JSON của tổ (VD: Tổ 1)...")
        self.btn_browse = QPushButton("📂 Chọn Folder")
        self.btn_browse.clicked.connect(self.browse_folder)
        folder_layout.addWidget(QLabel("📂 Folder Tổ:"))
        folder_layout.addWidget(self.folder_path)
        folder_layout.addWidget(self.btn_browse)
        self.folder_frame.setVisible(False)
        layout.addWidget(self.folder_frame)

        self.radio_single.toggled.connect(self.on_mode_toggled)



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

    def browse_folder(self):
        folder = QFileDialog.getExistingDirectory(self, "Chọn Folder Tổ")
        if folder:
            self.folder_path.setText(folder)

    def extract_dept_from_folder(self, folder):
        import json
        for f in os.listdir(folder):
            if f.endswith(".json"):
                try:
                    with open(os.path.join(folder, f), "r", encoding="utf-8") as file:
                        data = json.load(file)
                        to_sx = str(data.get("thong_tin_chung", {}).get("to_san_xuat", "")).strip()
                        if to_sx:
                            return f"Tổ {to_sx}" if to_sx.isdigit() else to_sx
                except Exception:
                    pass
        return None

    def on_mode_toggled(self, checked):
        self.folder_frame.setVisible(checked)
        self.date_frame.setDisabled(checked)

    def start_processing(self):
        # Bắt đầu gom JSON và đẩy vào Excel
        self.log_area.clear()
        self.progress_bar.setValue(0)
        
        month = self.combo_month.currentText()
        year = self.combo_year.currentText()
        single_dept = None

        if self.radio_single.isChecked():
            folder = self.folder_path.text().strip()
            if not folder or not os.path.exists(folder):
                QMessageBox.warning(self, "Lỗi", "Vui lòng chọn folder hợp lệ cho tổ cần chạy.")
                return
            
            # Tự động trích xuất month và year từ đường dẫn thư mục được chọn
            # Ví dụ: D:/Linh_Salary_Tool/02_ma_hang/2026/04/Tổ 1
            norm_path = folder.replace("\\", "/")
            parts = [p.strip() for p in norm_path.split("/") if p.strip()]
            
            if len(parts) >= 3:
                potential_month = parts[-2]
                potential_year = parts[-3]
                if potential_month.isdigit() and len(potential_month) == 2 and potential_year.isdigit() and len(potential_year) == 4:
                    month = potential_month
                    year = potential_year
            
            # Thử lấy tên tổ từ JSON trong folder
            single_dept = self.extract_dept_from_folder(folder)
            if not single_dept:
                # Nếu không có JSON, thử lấy tên folder được chọn làm tên Tổ (Ví dụ: "Tổ 3")
                last_dir = parts[-1]
                if "tổ" in last_dir.lower() or last_dir.isdigit():
                    single_dept = f"Tổ {last_dir}" if last_dir.isdigit() else last_dir
            
            if not single_dept:
                QMessageBox.warning(self, "Lỗi", "Không thể xác định tên Tổ từ thư mục đã chọn.")
                return

        json_root = os.path.join("D:/Linh_Salary_Tool", "02_ma_hang", year, month)
        excel_dir = get_split_excel_path(year, month)

        if not os.path.exists(json_root):
            QMessageBox.warning(self, "Lỗi", f"Không tìm thấy dữ liệu JSON của tháng {month}/{year} tại đường dẫn: {json_root}")
            return

        # Lấy mật khẩu từ file .env
        from dotenv import load_dotenv
        load_dotenv()
        env_pwd = os.getenv("EXCEL_SHEET_PASSWORD", "8863")

        self.set_controls_enabled(False)
        self.thread = IntegrationThread(json_root, excel_dir, env_pwd, single_dept)
        self.thread.log_signal.connect(self.log_area.append)
        self.thread.progress_signal.connect(self.progress_bar.setValue)
        self.thread.finished_signal.connect(self.on_finished)
        self.thread.start()

    def set_controls_enabled(self, enabled):
        self.mode_frame.setEnabled(enabled)
        self.folder_frame.setEnabled(enabled)
        if enabled:
            self.date_frame.setEnabled(self.radio_all.isChecked())
        else:
            self.date_frame.setEnabled(False)
        self.btn_run.setEnabled(enabled)

    def on_finished(self, success, message):
        # Kết thúc tích hợp
        self.set_controls_enabled(True)
        if success:
            self.progress_bar.setValue(100)
            QMessageBox.information(self, "Xong", message)
        else:
            QMessageBox.critical(self, "Lỗi", message)
