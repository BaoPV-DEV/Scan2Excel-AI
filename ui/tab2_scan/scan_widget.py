import os
import json
import re
from datetime import datetime
from PySide6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QPushButton, QLabel, QProgressBar, QTextEdit, 
    QFileDialog, QScrollArea, QSplitter, QLineEdit, QMessageBox, QComboBox, QInputDialog, QFrame, QDialog
)
from PySide6.QtCore import Qt
from PySide6.QtGui import QPixmap, QFont, QIntValidator
from dotenv import load_dotenv

# Import các thành phần nội bộ từ package tab2_scan
from ui.tab2_scan.drop_label import DropLabel
from ui.tab2_scan.scan_thread import ProcessingThread
from ui.tab2_scan.data_dialog import DataEditorDialog
from utils.paths import get_json_data_path

# Tải cấu hình ban đầu
load_dotenv()
INITIAL_API_KEY = os.getenv("GEMINI_API_KEY", "")

# Class Giao diện chính cho chức năng Scan Ảnh Sản Lượng (Tab 2).
class ScanProductionWidget(QWidget):
    # Khởi tạo và thiết lập Widget
    def __init__(self):
        super().__init__()
        self.setAcceptDrops(True)
        self.image_paths = []
        self.processed_data = None
        self.current_api_key = INITIAL_API_KEY
        
        self.setStyleSheet("""
            QPushButton { background-color: #3B82F6; color: white; border: none; border-radius: 6px; padding: 10px 15px; font-size: 14px; font-weight: bold; }
            QPushButton:hover { background-color: #2563EB; }
            QPushButton#ProcessBtn { background-color: #10B981; font-size: 15px; }
            QPushButton#ProcessBtn:hover { background-color: #059669; }
            QPushButton#ClearBtn { background-color: #EF4444; }
            QPushButton#UnlockBtn { background-color: #6B7280; font-size: 12px; padding: 5px 10px; }
            QProgressBar { border: 1px solid #E5E7EB; border-radius: 5px; text-align: center; background-color: #F3F4F6; height: 20px; }
            QProgressBar::chunk { background-color: #10B981; border-radius: 4px; }
            QTextEdit { background-color: #1E1E1E; color: #10B981; font-family: Consolas; font-size: 13px; border-radius: 6px; padding: 10px; }
            QLineEdit:disabled { background-color: #E5E7EB; color: #6B7280; }
        """)
        self.init_ui()
        
    # Thiết lập bố cục giao diện
    def init_ui(self):
        main_layout = QHBoxLayout(self)
        main_layout.setContentsMargins(15, 15, 15, 15)
        main_layout.setSpacing(20)
        
        splitter = QSplitter(Qt.Horizontal)
        
        # Cột trái: Upload và Preview ảnh
        left_panel = QWidget()
        left_layout = QVBoxLayout(left_panel)
        self.upload_btn = QPushButton("📁 Tải Ảnh Lên")
        self.upload_btn.clicked.connect(self.upload_image)
        self.drop_label = DropLabel("Kéo thả ảnh vào đây\n\n(Hỗ trợ PNG, JPG, JPEG)")
        self.preview_label = QLabel("Chưa có ảnh nào được tải")
        self.preview_label.setAlignment(Qt.AlignCenter)
        self.preview_label.setStyleSheet("background-color: #E5E7EB; border-radius: 8px;")
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setWidget(self.preview_label)
        scroll.setStyleSheet("border: none;")
        
        left_layout.addWidget(self.upload_btn)
        left_layout.addWidget(self.drop_label)
        left_layout.addWidget(QLabel("🖼️ Xem Trước Ảnh", font=QFont("Arial", 12, QFont.Bold)))
        left_layout.addWidget(scroll, stretch=1)
        
        # Cột phải: Cấu hình và Nhật ký
        right_panel = QWidget()
        right_layout = QVBoxLayout(right_panel)
        
        # --- Phần Cấu hình Model và Key ---
        config_group = QFrame()
        config_group.setStyleSheet("QFrame { background-color: #F3F4F6; border-radius: 8px; padding: 10px; }")
        config_layout = QVBoxLayout(config_group)
        
        # Model Selection
        model_layout = QHBoxLayout()
        model_layout.addWidget(QLabel("🤖 Model AI:"))
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
        self.model_combo.setStyleSheet("padding: 5px; background: white;")
        model_layout.addWidget(self.model_combo, 1)
        config_layout.addLayout(model_layout)
        
        # API Key Input
        key_layout = QHBoxLayout()
        key_layout.addWidget(QLabel("🔑 Gemini Key:"))
        self.key_input = QLineEdit(self.current_api_key)
        self.key_input.setEchoMode(QLineEdit.Password)
        self.key_input.setDisabled(True)
        self.key_input.setStyleSheet("padding: 5px;")
        key_layout.addWidget(self.key_input, 1)
        
        self.unlock_btn = QPushButton("🔓 Đổi Key")
        self.unlock_btn.setObjectName("UnlockBtn")
        self.unlock_btn.clicked.connect(self.unlock_key_input)
        key_layout.addWidget(self.unlock_btn)
        config_layout.addLayout(key_layout)
        
        right_layout.addWidget(QLabel("⚙️ Cấu Hình Hệ Thống", font=QFont("Arial", 12, QFont.Bold)))
        right_layout.addWidget(config_group)
        
        # Nhập mức thưởng
        bonus_layout = QHBoxLayout()
        bonus_layout.addWidget(QLabel("💰 Thưởng mã hàng mới (%):"))
        self.bonus_input = QLineEdit("0")
        self.bonus_input.setValidator(QIntValidator(0, 999999999))
        self.bonus_input.setStyleSheet("padding: 8px; font-size: 14px;")
        bonus_layout.addWidget(self.bonus_input)
        right_layout.addLayout(bonus_layout)
        
        # Nút bắt đầu OCR
        self.process_btn = QPushButton("🚀 BẮT ĐẦU QUÉT ẢNH (AI)")
        self.process_btn.setObjectName("ProcessBtn")
        self.process_btn.setEnabled(False)
        self.process_btn.clicked.connect(self.start_processing)
        self.progress_bar = QProgressBar()
        
        # Nhóm các nút chức năng phụ
        actions_layout = QHBoxLayout()
        self.preview_btn = QPushButton("👁️ Xem / Sửa Dữ Liệu")
        self.preview_btn.setEnabled(False)
        self.preview_btn.clicked.connect(self.preview_json)
        
        self.import_json_btn = QPushButton("📂 Mở File JSON")
        self.import_json_btn.clicked.connect(self.import_json)
        self.import_json_btn.setStyleSheet("background-color: #F59E0B;") 
        
        self.clear_btn = QPushButton("🗑️ Xóa")
        self.clear_btn.setObjectName("ClearBtn")
        self.clear_btn.clicked.connect(self.clear_data)
        
        actions_layout.addWidget(self.preview_btn)
        actions_layout.addWidget(self.import_json_btn)
        actions_layout.addWidget(self.clear_btn)
        
        self.log_panel = QTextEdit()
        self.log_panel.setReadOnly(True)
        
        right_layout.addWidget(self.process_btn)
        right_layout.addWidget(self.progress_bar)
        right_layout.addLayout(actions_layout)
        right_layout.addWidget(QLabel("📝 Nhật Ký Xử Lý", font=QFont("Arial", 12, QFont.Bold)))
        right_layout.addWidget(self.log_panel, stretch=1)
        
        splitter.addWidget(left_panel)
        splitter.addWidget(right_panel)
        splitter.setSizes([400, 500])
        main_layout.addWidget(splitter)

    # Yêu cầu mật khẩu để đổi API Key
    def unlock_key_input(self):
        load_dotenv(override=True) # Refresh lại các biến từ .env
        admin_pass = os.getenv("ADMIN_PASSWORD", "09032001")
        
        password, ok = QInputDialog.getText(self, "Xác thực", "Nhập mật khẩu quản trị để đổi Key:", QLineEdit.Password)
        if ok and password == admin_pass: 
            # Nếu người dùng chưa gõ gì mới, hiển thị key mới nhất từ .env
            env_key = os.getenv("GEMINI_API_KEY", "")
            if not self.key_input.text() or self.key_input.text() == self.current_api_key:
                self.key_input.setText(env_key)
            
            self.key_input.setDisabled(False)
            self.key_input.setEchoMode(QLineEdit.Normal)
            self.unlock_btn.setText("🔒 Khóa lại")
            self.unlock_btn.clicked.disconnect()
            self.unlock_btn.clicked.connect(self.lock_key_input)
            self.log_panel.append("🔓 Đã mở khóa cấu hình API Key. Dữ liệu được cập nhật từ .env")
        elif ok:
            QMessageBox.warning(self, "Lỗi", "Mật khẩu không chính xác!")

    # Khóa lại ô nhập API Key và lưu giá trị
    def lock_key_input(self):
        self.current_api_key = self.key_input.text().strip()
        self.key_input.setDisabled(True)
        self.key_input.setEchoMode(QLineEdit.Password)
        self.unlock_btn.setText("🔓 Đổi Key")
        self.unlock_btn.clicked.disconnect()
        self.unlock_btn.clicked.connect(self.unlock_key_input)
        self.log_panel.append("🔒 Đã khóa và cập nhật API Key mới.")

    # Mở hộp thoại chọn file ảnh
    def upload_image(self):
        files, _ = QFileDialog.getOpenFileNames(self, "Chọn Ảnh Bảng Sản Lượng", "", "Images (*.png *.jpg *.jpeg)")
        if files: self.load_images(files)
        
    # Tự động tạo thư mục the def dragEnterEvent():
    def dragEnterEvent(self, e):
        if e.mimeData().hasUrls(): e.accept()
        
    # Tự động tạo thư mục the def dropEvent():
    def dropEvent(self, e):
        files = [u.toLocalFile() for u in e.mimeData().urls() if u.toLocalFile().lower().endswith(('.png', '.jpg', '.jpeg'))]
        if files: self.load_images(files)
        
    # Tự động tạo thư mục the def load_images():
    def load_images(self, paths):
        self.image_paths = paths
        self.preview_label.setPixmap(QPixmap(paths[0]).scaled(400, 400, Qt.KeepAspectRatio))
        self.process_btn.setEnabled(True)
        self.log_panel.append(f"🟢 Đã tải {len(paths)} ảnh.")
        
    # Khởi động luồng xử lý AI với model và key đã chọn
    def start_processing(self):
        api_key = self.key_input.text().strip() or self.current_api_key
        model_name = self.model_combo.currentText()

        if not api_key:
            self.log_panel.append("❌ Lỗi: Chưa cấu hình API Key.")
            return

        self.process_btn.setEnabled(False)
        self.log_panel.clear()
        
        # Khởi tạo luồng xử lý Gemini
        self.thread = ProcessingThread(self.image_paths, api_key, model_name)
        self.thread.progress.connect(self.progress_bar.setValue)
        self.thread.log.connect(self.log_panel.append)
        self.thread.finished.connect(self.on_finished)
        self.thread.start()
        
    # Xử lý kết quả trả về từ AI
    def on_finished(self, data):
        if not data:
            self.process_btn.setEnabled(True)
            return
            
        try:
            bonus_val = int(self.bonus_input.text() or "0")
            if "thong_tin_chung" in data:
                data["thong_tin_chung"]["thuong_ma_hang_moi"] = bonus_val
        except:
            pass
            
        self.processed_data = data
        self.preview_btn.setEnabled(True)
        self.process_btn.setEnabled(True)
        
        self.save_processed_data(data)
        self.log_panel.append("✅ Xong! Bạn có thể kiểm tra và sửa lại dữ liệu.")
        self.progress_bar.setValue(100)
        
    # Lưu dữ liệu JSON xuống ổ đĩa
    def save_processed_data(self, data):
        try:
            thong_tin = data.get("thong_tin_chung", {})
            thoi_gian = thong_tin.get("thoi_gian", "N/A")
            to_sx = thong_tin.get("to_san_xuat", "N/A")
            ma_hang = thong_tin.get("ma_hang", "processed_data")
            
            match_date = re.search(r'(\d{1,2})[/-](\d{4})', thoi_gian)
            month = match_date.group(1).zfill(2) if match_date else datetime.now().strftime("%m")
            year = match_date.group(2) if match_date else datetime.now().strftime("%Y")
            
            to_val = str(to_sx).strip()
            team_num_match = re.search(r'\d+', to_val)
            team_folder = f"Tổ {team_num_match.group()}" if team_num_match else "Tổ Unknown"
            
            json_dir = get_json_data_path(year, month, team_folder)
            c_ma_hang = re.sub(r'[\\/*?:"<>|]', "_", str(ma_hang)).strip() or "N_A"
            json_path = os.path.join(json_dir, f"{c_ma_hang}.json")
            
            with open(json_path, 'w', encoding='utf-8') as f:
                json.dump(data, f, indent=4, ensure_ascii=False)
            self.log_panel.append(f"💾 Đã lưu tại: {json_path}")
        except Exception as e:
            self.log_panel.append(f"⚠️ Lỗi lưu JSON: {e}")

    # Mở hộp thoại để xem và chỉnh sửa dữ liệu JSON trực tiếp
    def preview_json(self):
        if not self.processed_data: return
        dialog = DataEditorDialog(self.processed_data, self)
        if dialog.exec() == QDialog.Accepted:
            updated = dialog.get_updated_data()
            if updated:
                self.processed_data = updated
                self.save_processed_data(updated)
                self.log_panel.append("✅ Dữ liệu đã được cập nhật thành công.")
                
    # Mở file JSON đã có từ trước để kiểm tra hoặc sửa lại
    def import_json(self):
        file_path, _ = QFileDialog.getOpenFileName(self, "Chọn file JSON sản lượng", "", "JSON Files (*.json)")
        if not file_path: return
        
        try:
            with open(file_path, 'r', encoding='utf-8') as f:
                data = json.load(f)
                self.processed_data = data
                self.preview_btn.setEnabled(True)
                self.log_panel.append(f"✅ Đã tải dữ liệu từ file: {os.path.basename(file_path)}")
                self.preview_json() 
        except Exception as e:
            QMessageBox.critical(self, "Lỗi", f"Không thể đọc file JSON: {str(e)}")
            
    # Xóa trắng toàn bộ dữ liệu hiện tại trên giao diện
    def clear_data(self):
        self.image_paths = []
        self.processed_data = None
        self.preview_label.clear()
        self.preview_label.setText("Chưa có ảnh nào được tải")
        self.log_panel.clear()
        self.progress_bar.setValue(0)
        self.preview_btn.setEnabled(False)
        self.process_btn.setEnabled(False)
