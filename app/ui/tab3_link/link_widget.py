import os
import json
from datetime import datetime

from PySide6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QLabel, QLineEdit,
    QPushButton, QComboBox, QTextEdit, QFileDialog, QMessageBox, QFrame, QProgressBar, QRadioButton
)
from PySide6.QtCore import Qt
from PySide6.QtGui import QFont

from app.ui.tab3_link.link_thread import IntegrationThread
from app.utils.paths import get_json_data_path, get_split_excel_path, get_base_path

from dotenv import load_dotenv
load_dotenv()

_ENV_PASSWORD = os.getenv("EXCEL_SHEET_PASSWORD", "8863")


class LinkExcelWidget(QWidget):
    """
    Giao diện người dùng cho chức năng Đẩy Dữ Liệu vào Excel (Tab 3).
    """
    def __init__(self):
        super().__init__()

        self.setObjectName("LinkWidget")
        self.integration_thread = None

        self.setStyleSheet("""
            QWidget#LinkWidget { background-color: white; border-radius: 8px; }
            QLabel { font-size: 13px; font-weight: bold; color: #374151; }
            QLineEdit { padding: 8px; border: 1px solid #D1D5DB; border-radius: 5px; }
            QPushButton { background-color: #3B82F6; color: white; padding: 8px 15px; border-radius: 6px; font-weight: bold; }
            QPushButton#ProcessBtn { background-color: #10B981; font-size: 15px; padding: 12px; }
            QProgressBar { border: 1px solid #E5E7EB; border-radius: 5px; text-align: center; height: 20px; }
            QTextEdit { background-color: #1E1E1E; color: #10B981; font-family: Consolas; padding: 10px; }
            
            /* ✅ CHỈ THÊM ĐÚNG ĐOẠN NÀY ĐỂ FIX LỖI POPUP KHÔNG LÀM XẤU COMBOBOX */
            QComboBox {
                combobox-popup: 0;
            }
            QComboBox QAbstractItemView { 
                min-width: 100px; /* Khống chế chiều rộng khung đổ xuống vừa vặn chữ */
            }
        """)

        self.init_ui()

    # ---------------- UI ----------------
    def init_ui(self):
        layout = QVBoxLayout(self)

        title = QLabel("🔗 ĐẨY DỮ LIỆU SẢN LƯỢNG VÀO EXCEL")
        title.setFont(QFont("Arial", 16, QFont.Bold))
        title.setAlignment(Qt.AlignCenter)
        layout.addWidget(title)

        self.mode_frame = QFrame()
        mode_layout = QHBoxLayout(self.mode_frame)

        self.radio_all = QRadioButton("Chạy tất cả các tổ")
        self.radio_single = QRadioButton("Chạy 1 tổ")
        self.radio_all.setChecked(True)

        mode_layout.addWidget(QLabel("⚙️ Chế độ:"))
        mode_layout.addWidget(self.radio_all)
        mode_layout.addWidget(self.radio_single)
        layout.addWidget(self.mode_frame)

        self.date_frame = QFrame()
        date_layout = QHBoxLayout(self.date_frame)

        now = datetime.now()

        self.combo_month = QComboBox()
        self.combo_month.addItems([f"{i:02d}" for i in range(1, 13)])
        self.combo_month.setCurrentText(f"{now.month:02d}")

        self.combo_year = QComboBox()
        self.combo_year.addItems([str(y) for y in range(2024, 2031)])
        self.combo_year.setCurrentText(str(now.year))

        date_layout.addWidget(QLabel("📅 Chọn Tháng/Năm (Gốc quét):"))
        date_layout.addWidget(self.combo_month)
        date_layout.addWidget(self.combo_year)

        layout.addWidget(self.date_frame)

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

        self.btn_run = QPushButton("🚀 BẮT ĐẦU TÍCH HỢP")
        self.btn_run.setObjectName("ProcessBtn")
        self.btn_run.clicked.connect(self.start_processing)
        layout.addWidget(self.btn_run)

        self.progress_bar = QProgressBar()
        layout.addWidget(self.progress_bar)

        self.log_area = QTextEdit()
        self.log_area.setReadOnly(True)
        layout.addWidget(self.log_area)

    # ---------------- Helpers ----------------
    def browse_folder(self):
        folder = QFileDialog.getExistingDirectory(self, "Chọn Folder Tổ")
        if folder:
            self.folder_path.setText(folder)

    def extract_dept_from_folder(self, folder: str):
        try:
            for f in os.listdir(folder):
                if not f.endswith(".json"):
                    continue

                path = os.path.join(folder, f)
                with open(path, "r", encoding="utf-8") as file:
                    data = json.load(file)

                to_sx = str(data.get("thong_tin_chung", {}).get("to_san_xuat", "")).strip()
                if to_sx:
                    return f"Tổ {to_sx}" if to_sx.isdigit() else to_sx
        except Exception:
            return None
        return None

    def on_mode_toggled(self, checked):
        self.folder_frame.setVisible(checked)
        # SỬA LỖI: Khi tích chọn "Chạy 1 tổ", chúng ta vẫn cho phép hiển thị Date_frame để người dùng xác nhận tháng năm làm mốc chuẩn phòng hờ tự động bóc tách đường dẫn lỗi.
        # Hoặc giữ nguyên logic cũ tùy bạn, nhưng không nên khóa cứng hoàn toàn.

    # ---------------- Core logic ----------------
    def _resolve_folder_info(self, folder: str):
        """
        SỬA LỖI LOGIC: Hàm phân tích thông tin từ thư mục được lựa chọn một cách an toàn.
        """
        norm = folder.replace("\\", "/")
        parts = [p for p in norm.split("/") if p]

        # Khởi tạo mặc định lấy từ UI Combobox trước để làm phương án dự phòng an toàn
        month = self.combo_month.currentText()
        year = self.combo_year.currentText()
        dept = None

        # Thử nghiệm bóc tách tháng năm từ chuỗi đường dẫn thư mục
        if len(parts) >= 3:
            potential_month = parts[-2]
            potential_year = parts[-3]
            if potential_month.isdigit() and len(potential_month) == 2 and potential_year.isdigit() and len(potential_year) == 4:
                month = potential_month
                year = potential_year

        # Trích xuất phòng ban dựa vào dữ liệu tệp cấu trúc JSON bên trong
        dept = self.extract_dept_from_folder(folder)

        # Phương án dự phòng nếu các tệp JSON không ghi trường dữ liệu "to_san_xuat"
        if not dept and parts:
            last = parts[-1]
            if "tổ" in last.lower():
                dept = last
            elif last.isdigit():
                dept = f"Tổ {last}"

        return month, year, dept

    def start_processing(self):
        if self.integration_thread and self.integration_thread.isRunning():
            return

        self.log_area.clear()
        self.progress_bar.setValue(0)

        # Lấy giá trị cơ sở ban đầu từ UI
        month = self.combo_month.currentText()
        year = self.combo_year.currentText()
        single_dept = None

        if self.radio_single.isChecked():
            folder = self.folder_path.text().strip()

            if not folder or not os.path.exists(folder):
                QMessageBox.warning(self, "Lỗi", "Vui lòng chọn folder hợp lệ cho tổ cần chạy.")
                return

            # Gọi bộ giải mã thông tin đường dẫn thư mục
            month, year, single_dept = self._resolve_folder_info(folder)

            if not single_dept:
                QMessageBox.warning(self, "Lỗi", "Không thể xác định tên Tổ từ thư mục đã chọn.")
                return

       # 1. Đường dẫn mặc định (Chạy toàn bộ các tổ của tháng/năm)
        json_root = os.path.join(get_base_path(), "02_ma_hang", year, month)
        excel_dir = get_split_excel_path(year, month) + "/sx"
        
        # 2. Xử lý đường dẫn nếu người dùng chọn chạy đơn lẻ 1 tổ từ giao diện
        if self.radio_single.isChecked():
            folder_selected = self.folder_path.text().strip()
            
            # Nếu người dùng chọn một folder cụ thể và folder đó tồn tại
            if folder_selected and os.path.exists(folder_selected):
                # Gán thẳng folder_selected làm gốc quét để os.walk() chạy nhanh nhất trong chính tổ đó
                json_root = folder_selected
                
                # Trích xuất tự động tên tổ từ chính tên thư mục được chọn để làm single_dept (nếu cần)
                if not single_dept:
                    single_dept = os.path.basename(folder_selected.replace("\\", "/"))

        print(f"DEBUG: json_root={json_root}, excel_dir={excel_dir}, month={month}, year={year}, single_dept={single_dept}")

        if not os.path.exists(json_root):
            QMessageBox.warning(
                self,
                "Lỗi",
                f"Không tìm thấy dữ liệu JSON tại đường dẫn yêu cầu: {json_root}"
            )
            return

        if not _ENV_PASSWORD:
            QMessageBox.warning(self, "Lỗi", "Thiếu EXCEL_SHEET_PASSWORD trong file .env")
            return

        self.set_controls_enabled(False)
        self.log_area.append(f"⏳ Đang chuẩn bị luồng xử lý dữ liệu cho Tháng {month}/{year}...")

        # Kích hoạt Thread xử lý ngầm chống đứng đóng băng ứng dụng
        self.integration_thread = IntegrationThread(
            json_root,
            excel_dir,
            _ENV_PASSWORD,
            single_dept
        )

        self.integration_thread.log_signal.connect(self.log_area.append)
        self.integration_thread.progress_signal.connect(self.progress_bar.setValue)
        self.integration_thread.finished_signal.connect(self.on_finished)
        self.integration_thread.start()

    # ---------------- UI state ----------------
    def set_controls_enabled(self, enabled: bool):
        self.mode_frame.setEnabled(enabled)
        self.folder_frame.setEnabled(enabled)
        self.date_frame.setEnabled(enabled and self.radio_all.isChecked())
        self.btn_run.setEnabled(enabled)

    def on_finished(self, success: bool, message: str):
        self.set_controls_enabled(True)

        if success:
            self.progress_bar.setValue(100)
            QMessageBox.information(self, "Xong", message)
        else:
            QMessageBox.critical(self, "Lỗi", message)