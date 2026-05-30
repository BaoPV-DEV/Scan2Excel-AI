import os
import json
import re
from datetime import datetime
from PySide6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QPushButton, QLabel, QProgressBar, QTextEdit, 
    QFileDialog, QScrollArea, QSplitter, QLineEdit, QMessageBox, QComboBox, QInputDialog, QFrame, QDialog,
    QTableWidget, QTableWidgetItem, QHeaderView
)
from PySide6.QtCore import Qt
from PySide6.QtGui import QPixmap, QFont, QIntValidator
from dotenv import load_dotenv

# Import các thành phần nội bộ từ package tab2_scan
from app.ui.tab2_scan.drop_label import DropLabel
from app.ui.tab2_scan.scan_thread import BatchProcessingThread
from app.ui.tab2_scan.data_dialog import DataEditorDialog
from app.utils.paths import get_json_data_path

# Tải cấu hình ban đầu
load_dotenv()
INITIAL_API_KEY = os.getenv("GEMINI_API_KEY", "")

class ScanProductionWidget(QWidget):
    """
    Giao diện Quét Ảnh Sản Lượng (Tab 2) - Nâng cấp hỗ trợ Batch OCR hàng loạt theo cấu trúc thư mục mã hàng.
    """
    def __init__(self):
        super().__init__()
        self.setAcceptDrops(True)
        self.batch_items = []
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
            QTableWidget { background-color: white; border: 1px solid #E5E7EB; border-radius: 6px; gridline-color: #F3F4F6; }
            QHeaderView::section { background-color: #F3F4F6; padding: 6px; border: 1px solid #E5E7EB; font-weight: bold; font-size: 12px; }
        """)
        self.init_ui()
        
    def init_ui(self):
        main_layout = QHBoxLayout(self)
        main_layout.setContentsMargins(15, 15, 15, 15)
        main_layout.setSpacing(20)
        
        splitter = QSplitter(Qt.Horizontal)
        
        # -------------------------------------------------------------
        # Cột trái: Cấu hình đầu vào & Quét thư mục
        # -------------------------------------------------------------
        left_panel = QWidget()
        left_layout = QVBoxLayout(left_panel)
        left_layout.setContentsMargins(0, 0, 0, 0)
        
        left_layout.addWidget(QLabel("📂 CẤU HÌNH & TẢI THƯ MỤC", font=QFont("Arial", 13, QFont.Bold)))
        
        self.upload_btn = QPushButton("📂 CHỌN THƯ MỤC QUÉT BATCH")
        self.upload_btn.clicked.connect(self.upload_folder)
        left_layout.addWidget(self.upload_btn)
        
        self.drop_label = DropLabel("Kéo thả thư mục chứa các mã hàng vào đây\n\n(Hỗ trợ thư mục chứa ảnh PNG, JPG, JPEG)")
        left_layout.addWidget(self.drop_label)
        
        self.folder_label = QLabel("Chưa chọn thư mục quét hàng loạt nào.")
        self.folder_label.setStyleSheet("color: #6B7280; font-style: italic; padding: 5px;")
        self.folder_label.setWordWrap(True)
        left_layout.addWidget(self.folder_label)
        
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
        
        left_layout.addWidget(QLabel("⚙️ Cấu Hình Hệ Thống", font=QFont("Arial", 12, QFont.Bold)))
        left_layout.addWidget(config_group)
        left_layout.addStretch(1)
        
        # -------------------------------------------------------------
        # Cột phải: Bảng danh sách Mã hàng Batch & Logs
        # -------------------------------------------------------------
        right_panel = QWidget()
        right_layout = QVBoxLayout(right_panel)
        right_layout.setContentsMargins(0, 0, 0, 0)
        
        right_layout.addWidget(QLabel("📋 DANH SÁCH MÃ HÀNG BATCH", font=QFont("Arial", 13, QFont.Bold)))
        
        # Bảng hiển thị danh sách Batch
        self.table_batch = QTableWidget()
        self.table_batch.setColumnCount(6)
        self.table_batch.setHorizontalHeaderLabels(["STT", "Thư mục con (Mã)", "Số ảnh", "Thưởng mã (%)", "Trạng thái", "Hành động"])
        
        header = self.table_batch.horizontalHeader()
        # header.setSectionResizeMode(0, QHeaderView.ResizeToContents)
        # header.setSectionResizeMode(1, QHeaderView.Stretch)
        # header.setSectionResizeMode(2, QHeaderView.ResizeToContents)
        # header.setSectionResizeMode(3, QHeaderView.ResizeToContents)
        # header.setSectionResizeMode(4, QHeaderView.ResizeToContents)
        # header.setSectionResizeMode(5, QHeaderView.ResizeToContents)
        # STT
        header.setSectionResizeMode(0, QHeaderView.Fixed)
        self.table_batch.setColumnWidth(0, 0)

        # Thư mục con
        header.setSectionResizeMode(1, QHeaderView.Stretch)

        # Số ảnh
        header.setSectionResizeMode(2, QHeaderView.Fixed)
        self.table_batch.setColumnWidth(2, 70)

        # Thưởng mã
        header.setSectionResizeMode(3, QHeaderView.Fixed)
        self.table_batch.setColumnWidth(3, 120)

        # Trạng thái
        header.setSectionResizeMode(4, QHeaderView.Fixed)
        self.table_batch.setColumnWidth(4, 100)

        # Hành động
        header.setSectionResizeMode(5, QHeaderView.Fixed)
        self.table_batch.setColumnWidth(5, 120)

        right_layout.addWidget(self.table_batch, stretch=3)
        
        # Nút bắt đầu Batch OCR
        self.process_btn = QPushButton("🚀 BẮT ĐẦU QUÉT HÀNG LOẠT (AI)")
        self.process_btn.setObjectName("ProcessBtn")
        self.process_btn.setEnabled(False)
        self.process_btn.clicked.connect(self.start_processing)
        
        self.progress_bar = QProgressBar()
        self.progress_bar.setValue(0)
        
        # Nhóm các nút chức năng phụ
        actions_layout = QHBoxLayout()
        self.import_json_btn = QPushButton("📂 Mở File JSON")
        self.import_json_btn.clicked.connect(self.import_json)
        self.import_json_btn.setStyleSheet("background-color: #F59E0B;") 
        
        self.clear_btn = QPushButton("🗑️ Xóa")
        self.clear_btn.setObjectName("ClearBtn")
        self.clear_btn.clicked.connect(self.clear_data)
        
        actions_layout.addWidget(self.import_json_btn)
        actions_layout.addWidget(self.clear_btn)
        
        self.log_panel = QTextEdit()
        self.log_panel.setReadOnly(True)
        
        right_layout.addWidget(self.process_btn)
        right_layout.addWidget(self.progress_bar)
        right_layout.addLayout(actions_layout)
        right_layout.addWidget(QLabel("📝 Nhật Ký Xử Lý", font=QFont("Arial", 12, QFont.Bold)))
        right_layout.addWidget(self.log_panel, stretch=2)
        
        splitter.addWidget(left_panel)
        splitter.addWidget(right_panel)
        splitter.setSizes([380, 620])
        main_layout.addWidget(splitter)

    def unlock_key_input(self):
        load_dotenv(override=True)
        admin_pass = os.getenv("ADMIN_PASSWORD", "09032001")
        password, ok = QInputDialog.getText(self, "Xác thực", "Nhập mật khẩu quản trị để đổi Key:", QLineEdit.Password)
        if ok and password == admin_pass: 
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

    def lock_key_input(self):
        self.current_api_key = self.key_input.text().strip()
        self.key_input.setDisabled(True)
        self.key_input.setEchoMode(QLineEdit.Password)
        self.unlock_btn.setText("🔓 Đổi Key")
        self.unlock_btn.clicked.disconnect()
        self.unlock_btn.clicked.connect(self.unlock_key_input)
        self.log_panel.append("🔒 Đã khóa và cập nhật API Key mới.")

    def upload_folder(self):
        folder = QFileDialog.getExistingDirectory(self, "Chọn Thư Mục Quét Hàng Loạt")
        if folder:
            self.folder_label.setText(f"📁 Đường dẫn: {folder}")
            self.folder_label.setToolTip(folder)
            self.load_batch_folders(folder)

    def dragEnterEvent(self, e):
        if e.mimeData().hasUrls():
            e.accept()

    def dropEvent(self, e):
        folders = [u.toLocalFile() for u in e.mimeData().urls() if os.path.isdir(u.toLocalFile())]
        if folders:
            folder = folders[0]
            self.folder_label.setText(f"📁 Đường dẫn: {folder}")
            self.folder_label.setToolTip(folder)
            self.load_batch_folders(folder)

    def load_batch_folders(self, root_folder):
        self.batch_items = []
        self.table_batch.setRowCount(0)
        
        try:
            subdirs = sorted([d for d in os.listdir(root_folder) if os.path.isdir(os.path.join(root_folder, d))])
        except Exception as e:
            self.log_panel.append(f"❌ Lỗi đọc thư mục: {e}")
            return
        
        row_idx = 0
        for subdir in subdirs:
            subdir_path = os.path.join(root_folder, subdir)
            images = []
            for f in os.listdir(subdir_path):
                if f.lower().endswith(('.png', '.jpg', '.jpeg')):
                    images.append(os.path.join(subdir_path, f))
            
            if not images:
                continue
            
            self.table_batch.insertRow(row_idx)
            
            # STT
            self.table_batch.setItem(row_idx, 0, QTableWidgetItem(str(row_idx + 1)))
            # Tên thư mục con (Mã)
            self.table_batch.setItem(row_idx, 1, QTableWidgetItem(subdir))
            # Số ảnh
            self.table_batch.setItem(row_idx, 2, QTableWidgetItem(f"{len(images)} ảnh"))
            
            # Thưởng mã hàng mới (%)
            bonus_edit = QLineEdit("0")
            bonus_edit.setValidator(QIntValidator(0, 100))
            bonus_edit.setAlignment(Qt.AlignCenter)
            bonus_edit.setStyleSheet("padding: 2px; border: 1px solid #D1D5DB; border-radius: 3px; min-width: 40px;")
            self.table_batch.setCellWidget(row_idx, 3, bonus_edit)
            
            # Trạng thái
            status_item = QTableWidgetItem("Chưa quét")
            status_item.setTextAlignment(Qt.AlignCenter)
            self.table_batch.setItem(row_idx, 4, status_item)
            
            # Hành động
            action_widget = QWidget()
            action_layout = QHBoxLayout(action_widget)
            action_layout.setContentsMargins(2, 2, 2, 2)
            action_layout.setSpacing(5)
            
            view_btn = QPushButton("👁️")
            view_btn.setToolTip("Xem/Sửa dữ liệu")
            view_btn.setStyleSheet("background-color: #3B82F6; color: white; padding: 2px; border-radius: 3px; min-width: 25px; font-size: 11px;")
            view_btn.setEnabled(False)
            
            retry_btn = QPushButton("🔄")
            retry_btn.setToolTip("Quét lại mã này")
            retry_btn.setStyleSheet("background-color: #F59E0B; color: white; padding: 2px; border-radius: 3px; min-width: 25px; font-size: 11px;")
            retry_btn.setEnabled(False)
            
            action_layout.addWidget(view_btn)
            action_layout.addWidget(retry_btn)
            action_layout.setAlignment(Qt.AlignCenter)
            
            self.table_batch.setCellWidget(row_idx, 5, action_widget)
            
            item_data = {
                "row": row_idx,
                "subdir_name": subdir,
                "subdir_path": subdir_path,
                "images": images,
                "bonus_edit": bonus_edit,
                "view_btn": view_btn,
                "retry_btn": retry_btn,
                "processed_json": None
            }
            self.batch_items.append(item_data)
            
            view_btn.clicked.connect(lambda checked=False, r=row_idx: self.view_single_batch_item(r))
            retry_btn.clicked.connect(lambda checked=False, r=row_idx: self.retry_single_batch_item(r))
            
            row_idx += 1
            
        self.process_btn.setEnabled(len(self.batch_items) > 0)
        self.log_panel.append(f"🟢 Đã load {len(self.batch_items)} mã hàng từ thư mục con.")

    def set_controls_enabled(self, enabled):
        self.upload_btn.setEnabled(enabled)
        self.model_combo.setEnabled(enabled)
        self.key_input.setEnabled(enabled and self.unlock_btn.text() == "🔒 Khóa lại")
        self.unlock_btn.setEnabled(enabled)
        self.process_btn.setEnabled(enabled and len(self.batch_items) > 0)
        self.import_json_btn.setEnabled(enabled)
        self.clear_btn.setEnabled(enabled)
        
        for item in self.batch_items:
            item["bonus_edit"].setEnabled(enabled)
            item["view_btn"].setEnabled(enabled and item["processed_json"] is not None)
            row = item["row"]
            status_text = self.table_batch.item(row, 4).text()
            item["retry_btn"].setEnabled(enabled and status_text != "Chưa quét")

    def start_processing(self):
        api_key = self.key_input.text().strip() or self.current_api_key
        if not api_key:
            QMessageBox.warning(self, "Lỗi", "Chưa cấu hình API Key.")
            return
        
        # Chỉ quét các mã hàng chưa quét thành công
        batch_data = []
        for item in self.batch_items:
            if item["processed_json"] is None:
                batch_data.append({
                    "row": item["row"],
                    "subdir_name": item["subdir_name"],
                    "subdir_path": item["subdir_path"],
                    "images": item["images"],
                    "bonus": int(item["bonus_edit"].text() or 0)
                })
        
        if not batch_data:
            QMessageBox.information(self, "Thông báo", "Tất cả các mã hàng đã được quét thành công!")
            return
        
        self.set_controls_enabled(False)
        self.progress_bar.setValue(0)
        self.log_panel.append(f"🚀 Khởi chạy quét hàng loạt {len(batch_data)} mã hàng...")
        
        model_name = self.model_combo.currentText()
        self.thread = BatchProcessingThread(batch_data, api_key, model_name)
        self.thread.item_started.connect(self.on_item_started)
        self.thread.item_finished.connect(self.on_item_finished)
        self.thread.progress.connect(self.progress_bar.setValue)
        self.thread.log.connect(self.log_panel.append)
        self.thread.finished.connect(self.on_batch_finished)
        self.thread.start()

    def on_item_started(self, row, subdir_name):
        status_item = self.table_batch.item(row, 4)
        status_item.setText("🟡 Đang quét...")
        status_item.setForeground(Qt.blue)

    def on_item_finished(self, row, status, data, msg_or_path):
        status_item = self.table_batch.item(row, 4)
        item = self.batch_items[row]
        
        if status == "success":
            status_item.setText("🟢 Thành công")
            status_item.setForeground(Qt.darkGreen)
            item["processed_json"] = data
            item["view_btn"].setEnabled(True)
            item["retry_btn"].setEnabled(True)
            self.log_panel.append(f"✔️ Đã lưu kết quả tại {os.path.basename(msg_or_path)}")
        else:
            status_item.setText("🔴 Lỗi")
            status_item.setForeground(Qt.red)
            item["retry_btn"].setEnabled(True)
            self.log_panel.append(f"❌ Mã hàng {item['subdir_name']} quét thất bại: {msg_or_path}")

    def on_batch_finished(self):
        self.set_controls_enabled(True)
        self.progress_bar.setValue(100)
        self.log_panel.append("🎉 Đã hoàn thành toàn bộ tiến trình quét Batch!")
        QMessageBox.information(self, "Hoàn tất", "Hoàn thành quét hàng loạt!")

    def retry_single_batch_item(self, row_idx):
        api_key = self.key_input.text().strip() or self.current_api_key
        if not api_key:
            QMessageBox.warning(self, "Lỗi", "Chưa cấu hình API Key.")
            return
            
        item = self.batch_items[row_idx]
        single_batch_data = [{
            "row": item["row"],
            "subdir_name": item["subdir_name"],
            "subdir_path": item["subdir_path"],
            "images": item["images"],
            "bonus": int(item["bonus_edit"].text() or 0)
        }]
        
        self.set_controls_enabled(False)
        self.log_panel.append(f"🔄 Quét lại mã riêng lẻ: {item['subdir_name']}...")
        
        model_name = self.model_combo.currentText()
        self.thread = BatchProcessingThread(single_batch_data, api_key, model_name)
        self.thread.item_started.connect(self.on_item_started)
        self.thread.item_finished.connect(self.on_item_finished)
        self.thread.log.connect(self.log_panel.append)
        self.thread.finished.connect(self.on_retry_finished)
        self.thread.start()

    def on_retry_finished(self):
        self.set_controls_enabled(True)
        self.log_panel.append("✨ Hoàn tất quét lại mã hàng riêng lẻ!")

    def view_single_batch_item(self, idx):
        item = self.batch_items[idx]
        if not item["processed_json"]:
            return
        
        dialog = DataEditorDialog(item["processed_json"], self)
        if dialog.exec() == QDialog.Accepted:
            updated = dialog.get_updated_data()
            if updated:
                item["processed_json"] = updated
                
                # Lưu lại JSON
                try:
                    thong_tin = updated.get("thong_tin_chung", {})
                    thoi_gian = thong_tin.get("thoi_gian", "N/A")
                    to_sx = thong_tin.get("to_san_xuat", "N/A")
                    ma_hang = thong_tin.get("ma_hang", item["subdir_name"])
                    
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
                        json.dump(updated, f, indent=4, ensure_ascii=False)
                    self.log_panel.append(f"💾 Đã cập nhật và lưu JSON: {json_path}")
                except Exception as e:
                    self.log_panel.append(f"⚠️ Lỗi lưu JSON cập nhật: {e}")

    def import_json(self):
        file_path, _ = QFileDialog.getOpenFileName(self, "Chọn file JSON sản lượng", "", "JSON Files (*.json)")
        if not file_path:
            return
        
        try:
            with open(file_path, 'r', encoding='utf-8') as f:
                data = json.load(f)
                
            dialog = DataEditorDialog(data, self)
            if dialog.exec() == QDialog.Accepted:
                updated = dialog.get_updated_data()
                if updated:
                    # Lưu lại file vừa chọn
                    with open(file_path, 'w', encoding='utf-8') as f:
                        json.dump(updated, f, indent=4, ensure_ascii=False)
                    self.log_panel.append(f"💾 Đã cập nhật và lưu trực tiếp file JSON: {os.path.basename(file_path)}")
        except Exception as e:
            QMessageBox.critical(self, "Lỗi", f"Không thể đọc/ghi file JSON: {str(e)}")

    def clear_data(self):
        self.batch_items = []
        self.table_batch.setRowCount(0)
        self.folder_label.setText("Chưa chọn thư mục quét hàng loạt nào.")
        self.log_panel.clear()
        self.progress_bar.setValue(0)
        self.process_btn.setEnabled(False)
