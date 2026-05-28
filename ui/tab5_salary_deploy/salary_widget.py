import os
import json
import openpyxl
from datetime import datetime
from PySide6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QLabel,
    QPushButton, QComboBox, QTextEdit, QMessageBox, QFrame, QProgressBar,
    QGroupBox, QLineEdit, QFileDialog, QScrollArea, QStackedWidget, QSizePolicy
)
from PySide6.QtCore import Qt, Signal
from PySide6.QtGui import QFont
from ui.tab5_salary_deploy.salary_thread import SalaryDeployThread


# Định nghĩa 5 nhóm file nguồn chính
SOURCE_FILE_DEFINITIONS = [
    {
        "key": "danh_sach_cbcnv_luong",
        "label": "📋 DS CBCNV bản dùng làm lương",
        "default_sheets": ["Danh sách"],
        "description": "File danh sách CBCNV bản dùng làm lương (VD: Danh sách CBCNV bản dùng làm lương T4.xlsx)"
    },
    {
        "key": "cham_cong",
        "label": "⏱️ Bảng chấm công tháng",
        "default_sheets": ["Công"],
        "description": "File chấm công tháng hiện tại (VD: Chấm công 04.26.xlsx)"
    },
    {
        "key": "chuyen_khoan_thang_truoc",
        "label": "💳 Chuyển khoản tháng trước",
        "default_sheets": ["Bản gốc T{prev_month}"],
        "description": "File CK tháng trước (VD: CK tháng 03.2026.xlsx)"
    },
    {
        "key": "danh_sach_cbcnv_thang",
        "label": "📝 DS CBCNV làm lương tháng",
        "default_sheets": ["Danh Sách dùng"],
        "description": "File danh sách CBCNV làm lương tháng (VD: Danh sách CBCNV làm lương tháng 04.xlsx)"
    },
    {
        "key": "phu_cap_con_nho",
        "label": "👶 Phụ cấp con nhỏ",
        "default_sheets": ["Con nhỏ", "Nhóm 6", "Thâm niên", "Sheet1"],
        "description": "File phụ cấp con nhỏ (VD: phụ cấp con nhỏ năm 2026.xlsx)"
    },
]

# Đường dẫn lưu cấu hình người dùng
def _get_user_config_path():
    from utils.paths import get_user_run_config_path
    return get_user_run_config_path()


class SourceFileRow(QFrame):
    """Widget cho 1 dòng cấu hình file nguồn."""
    config_changed = Signal()

    def __init__(self, definition: dict, parent=None):
        super().__init__(parent)
        self.definition = definition
        self.key = definition["key"]
        self.sheet_combos = {}  # sheet_pattern -> QComboBox
        self._init_ui()
    
    def _init_ui(self):
        self.setObjectName("SourceRow")
        self.setStyleSheet("""
            QFrame#SourceRow {
                border: 1px solid #E5E7EB; border-radius: 6px;
                background-color: #FAFBFC; padding: 8px; margin-bottom: 4px;
            }
        """)
        layout = QVBoxLayout(self)
        layout.setContentsMargins(10, 8, 10, 8)
        layout.setSpacing(6)
        
        # Dòng 1: Label + File path + Nút chọn
        row1 = QHBoxLayout()
        lbl = QLabel(self.definition["label"])
        lbl.setStyleSheet("font-weight: bold; font-size: 12px; color: #1F2937; min-width: 240px;")
        row1.addWidget(lbl)
        
        self.txt_path = QLineEdit()
        self.txt_path.setReadOnly(True)
        self.txt_path.setPlaceholderText("Chưa chọn file...")
        self.txt_path.setStyleSheet("""
            QLineEdit {
                padding: 5px 8px; border: 1px solid #D1D5DB; border-radius: 4px;
                background-color: #FFF; color: #374151; font-size: 12px;
            }
        """)
        row1.addWidget(self.txt_path, stretch=1)
        
        self.btn_browse = QPushButton("📂 Chọn file")
        self.btn_browse.setStyleSheet("""
            QPushButton {
                background-color: #3B82F6; color: white; border: none;
                border-radius: 4px; padding: 5px 12px; font-size: 11px; font-weight: bold;
            }
            QPushButton:hover { background-color: #2563EB; }
        """)
        self.btn_browse.setCursor(Qt.PointingHandCursor)
        self.btn_browse.clicked.connect(self._browse_file)
        row1.addWidget(self.btn_browse)
        
        self.btn_clear = QPushButton("✖")
        self.btn_clear.setStyleSheet("""
            QPushButton {
                background-color: #EF4444; color: white; border: none;
                border-radius: 4px; padding: 5px 8px; font-size: 11px; font-weight: bold;
            }
            QPushButton:hover { background-color: #DC2626; }
        """)
        self.btn_clear.setCursor(Qt.PointingHandCursor)
        self.btn_clear.clicked.connect(self._clear_selection)
        row1.addWidget(self.btn_clear)
        
        layout.addLayout(row1)
        
        # Dòng 2: Sheet dropdowns
        self.sheet_layout = QHBoxLayout()
        self.sheet_layout.setSpacing(10)
        
        for sheet_pattern in self.definition["default_sheets"]:
            sheet_lbl = QLabel(f"Sheet [{sheet_pattern}]:")
            sheet_lbl.setStyleSheet("font-size: 11px; color: #6B7280;")
            self.sheet_layout.addWidget(sheet_lbl)
            
            combo = QComboBox()
            combo.setMinimumWidth(160)
            combo.setStyleSheet("""
                QComboBox {
                    padding: 4px 8px; border: 1px solid #D1D5DB; border-radius: 4px;
                    background-color: #FFF; color: #374151; font-size: 11px;
                }
            """)
            combo.addItem(f"-- Mặc định: {sheet_pattern} --")
            self.sheet_layout.addWidget(combo)
            self.sheet_combos[sheet_pattern] = combo
        
        self.sheet_layout.addStretch()
        layout.addLayout(self.sheet_layout)
    
    def _browse_file(self):
        """Mở hộp thoại chọn file Excel"""
        file_path, _ = QFileDialog.getOpenFileName(
            self, f"Chọn file: {self.definition['label']}",
            "", "Excel Files (*.xlsx *.xls);;All Files (*)"
        )
        if file_path:
            self.txt_path.setText(file_path)
            self._load_sheets_from_file(file_path)
            self.config_changed.emit()

    def _clear_selection(self):
        """Xóa lựa chọn file"""
        self.txt_path.clear()
        for combo in self.sheet_combos.values():
            combo.clear()
            # Reset lại giá trị mặc định
        for sheet_pattern, combo in self.sheet_combos.items():
            combo.addItem(f"-- Mặc định: {sheet_pattern} --")
        self.config_changed.emit()

    def _load_sheets_from_file(self, file_path: str):
        """Đọc danh sách sheet từ file Excel và cập nhật các dropdown"""
        try:
            wb = openpyxl.load_workbook(file_path, read_only=True, data_only=True)
            sheet_names = wb.sheetnames
            wb.close()
            
            for sheet_pattern, combo in self.sheet_combos.items():
                combo.clear()
                combo.addItem(f"-- Mặc định: {sheet_pattern} --")
                for sn in sheet_names:
                    combo.addItem(sn)
                
                # Auto-select nếu có sheet trùng tên (case-insensitive)
                for i, sn in enumerate(sheet_names):
                    # So khớp pattern (loại bỏ placeholder)
                    clean_pattern = sheet_pattern.replace("{prev_month}", "").replace("{mm}", "").strip()
                    if sn.lower() == sheet_pattern.lower() or (clean_pattern and clean_pattern.lower() in sn.lower()):
                        combo.setCurrentIndex(i + 1)  # +1 vì có item mặc định ở đầu
                        break
        except Exception as e:
            QMessageBox.warning(self, "Lỗi đọc file", f"Không thể đọc danh sách sheet:\n{str(e)}")
    
    def get_config(self) -> dict:
        """Lấy cấu hình hiện tại của dòng này"""
        file_path = self.txt_path.text().strip()
        if not file_path:
            return {}
        
        sheet_mapping = {}
        for sheet_pattern, combo in self.sheet_combos.items():
            selected = combo.currentText()
            if selected and not selected.startswith("-- Mặc định"):
                sheet_mapping[sheet_pattern] = selected
        
        return {
            "file_path": file_path,
            "sheet_mapping": sheet_mapping
        }
    
    def set_config(self, cfg: dict):
        """Khôi phục cấu hình từ dict"""
        file_path = cfg.get("file_path", "")
        if file_path and os.path.exists(file_path):
            self.txt_path.setText(file_path)
            self._load_sheets_from_file(file_path)
            
            # Khôi phục sheet mapping
            sheet_mapping = cfg.get("sheet_mapping", {})
            for sheet_pattern, combo in self.sheet_combos.items():
                if sheet_pattern in sheet_mapping:
                    target_sheet = sheet_mapping[sheet_pattern]
                    idx = combo.findText(target_sheet)
                    if idx >= 0:
                        combo.setCurrentIndex(idx)
        self.config_changed.emit()


class SalaryDeployWidget(QWidget):
    """
    Giao diện Tab 5: Triển khai Công Lương (Salary Deployment).

    Giao diện 2 giai đoạn (QStackedWidget):
    - Giai đoạn 1: Cấu hình file nguồn (cuộn được, chiếm toàn màn hình tab)
    - Giai đoạn 2: Triển khai + nhật ký (ẩn config, log chiếm phần lớn diện tích)
    """
    PAGE_CONFIG = 0
    PAGE_DEPLOY = 1

    def __init__(self):
        super().__init__()
        self.setObjectName("SalaryWidget")
        self.thread = None
        self.source_rows = {}  # key -> SourceFileRow
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
            QGroupBox {
                font-weight: bold; font-size: 14px; color: #1F2937;
                border: 2px solid #3B82F6; border-radius: 8px;
                margin-top: 10px; padding-top: 20px;
            }
            QGroupBox::title {
                subcontrol-origin: margin; subcontrol-position: top left;
                padding: 2px 10px; background-color: #3B82F6; color: white;
                border-radius: 4px;
            }
        """)
        self.init_ui()
        self._load_user_config()
        self._update_deploy_ui()
        self._update_step_label()
        # Đã có cấu hình đủ từ lần trước → vào thẳng màn hình triển khai (log rộng)
        if self._is_source_config_complete()[0]:
            self._go_to_deploy_page()

    def _build_date_selector(self) -> QFrame:
        """Khối chọn tháng/năm dùng chung cho cả 2 giai đoạn."""
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
        return date_frame

    def _build_config_page(self) -> QWidget:
        """Giai đoạn 1: cấu hình file nguồn (cuộn được)."""
        page = QWidget()
        layout = QVBoxLayout(page)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(8)

        hint = QLabel(
            "💡 Chọn file Excel → Hệ thống đọc Sheet → Chọn Sheet phù hợp. "
            "Khi đủ 5 file, nhấn «Tiếp tục» để sang màn hình triển khai."
        )
        hint.setWordWrap(True)
        hint.setStyleSheet("color: #6B7280; font-size: 11px; font-weight: normal;")
        layout.addWidget(hint)

        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(QFrame.NoFrame)
        scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        scroll.setStyleSheet("QScrollArea { background: transparent; border: none; }")

        scroll_content = QWidget()
        scroll_layout = QVBoxLayout(scroll_content)
        scroll_layout.setContentsMargins(2, 2, 8, 2)
        scroll_layout.setSpacing(6)

        config_group = QGroupBox("⚙️ Cấu hình 5 file liên kết nguồn")
        config_layout = QVBoxLayout(config_group)
        config_layout.setSpacing(6)

        for defn in SOURCE_FILE_DEFINITIONS:
            row = SourceFileRow(defn, self)
            row.config_changed.connect(self._update_deploy_ui)
            config_layout.addWidget(row)
            self.source_rows[defn["key"]] = row

        scroll_layout.addWidget(config_group)
        scroll_layout.addStretch()
        scroll.setWidget(scroll_content)
        layout.addWidget(scroll, stretch=1)

        self.lbl_config_status = QLabel()
        self.lbl_config_status.setWordWrap(True)
        self.lbl_config_status.setAlignment(Qt.AlignCenter)
        self.lbl_config_status.setStyleSheet(
            "color: #B45309; font-size: 12px; font-weight: normal; "
            "background-color: #FFFBEB; border: 1px solid #FCD34D; "
            "border-radius: 6px; padding: 10px;"
        )
        layout.addWidget(self.lbl_config_status)

        self.btn_continue = QPushButton("✅ Tiếp tục — Sang màn hình triển khai")
        self.btn_continue.setObjectName("ProcessBtn")
        self.btn_continue.setCursor(Qt.PointingHandCursor)
        self.btn_continue.setVisible(False)
        self.btn_continue.clicked.connect(self._go_to_deploy_page)
        layout.addWidget(self.btn_continue)

        return page

    def _build_deploy_page(self) -> QWidget:
        """Giai đoạn 2: triển khai + nhật ký (log chiếm phần lớn)."""
        page = QWidget()
        layout = QVBoxLayout(page)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(8)

        summary_bar = QFrame()
        summary_bar.setObjectName("SectionFrame")
        bar_layout = QHBoxLayout(summary_bar)
        bar_layout.setContentsMargins(12, 8, 12, 8)

        self.lbl_deploy_summary = QLabel()
        self.lbl_deploy_summary.setStyleSheet("font-size: 12px; font-weight: normal; color: #374151;")
        bar_layout.addWidget(self.lbl_deploy_summary, stretch=1)

        self.btn_edit_config = QPushButton("⚙️ Sửa cấu hình nguồn")
        self.btn_edit_config.setStyleSheet("""
            QPushButton {
                background-color: #F3F4F6; color: #374151; border: 1px solid #D1D5DB;
                border-radius: 6px; padding: 6px 12px; font-size: 12px; font-weight: bold;
            }
            QPushButton:hover { background-color: #E5E7EB; }
        """)
        self.btn_edit_config.setCursor(Qt.PointingHandCursor)
        self.btn_edit_config.clicked.connect(self._go_to_config_page)
        bar_layout.addWidget(self.btn_edit_config)
        layout.addWidget(summary_bar)

        self.btn_run = QPushButton("🚀 BẮT ĐẦU TRIỂN KHAI CÔNG LƯƠNG")
        self.btn_run.setObjectName("ProcessBtn")
        self.btn_run.clicked.connect(self.start_processing)
        self.btn_run.setCursor(Qt.PointingHandCursor)
        layout.addWidget(self.btn_run)

        self.progress_bar = QProgressBar()
        self.progress_bar.setVisible(False)
        layout.addWidget(self.progress_bar)

        log_label = QLabel("📝 Nhật ký hoạt động")
        log_label.setStyleSheet("color: #374151; font-weight: bold; font-size: 13px;")
        layout.addWidget(log_label)

        self.log_area = QTextEdit()
        self.log_area.setReadOnly(True)
        self.log_area.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        self.log_area.setMinimumHeight(200)
        layout.addWidget(self.log_area, stretch=1)

        return page

    def init_ui(self):
        root = QVBoxLayout(self)
        root.setContentsMargins(20, 12, 20, 12)
        root.setSpacing(8)

        title = QLabel("💰 TRIỂN KHAI CÔNG THỨC LƯƠNG")
        title.setFont(QFont("Arial", 18, QFont.Bold))
        title.setAlignment(Qt.AlignCenter)
        title.setStyleSheet("color: #1F2937;")
        root.addWidget(title)

        self.lbl_step = QLabel()
        self.lbl_step.setAlignment(Qt.AlignCenter)
        self.lbl_step.setStyleSheet("color: #6B7280; font-size: 12px; font-weight: normal;")
        root.addWidget(self.lbl_step)

        root.addWidget(self._build_date_selector())

        self.stack = QStackedWidget()
        self.page_config = self._build_config_page()
        self.page_deploy = self._build_deploy_page()
        self.stack.addWidget(self.page_config)
        self.stack.addWidget(self.page_deploy)
        root.addWidget(self.stack, stretch=1)

    def _is_source_config_complete(self) -> tuple:
        """Kiểm tra đã chọn đủ và hợp lệ 5 file nguồn chưa."""
        missing = []
        for defn in SOURCE_FILE_DEFINITIONS:
            key = defn["key"]
            cfg = self.source_rows[key].get_config()
            path = cfg.get("file_path", "")
            if not path or not os.path.exists(path):
                missing.append(defn["label"])
        return len(missing) == 0, missing

    def _period_label(self) -> str:
        return f"Tháng {self.combo_month.currentText()}/{self.combo_year.currentText()}"

    def _refresh_deploy_summary(self):
        self.lbl_deploy_summary.setText(
            f"📅 Kỳ lương: {self._period_label()}  |  ✅ Đã cấu hình {len(SOURCE_FILE_DEFINITIONS)} file nguồn"
        )

    def _go_to_config_page(self):
        self.stack.setCurrentIndex(self.PAGE_CONFIG)
        self._update_step_label()

    def _go_to_deploy_page(self):
        complete, missing = self._is_source_config_complete()
        if not complete:
            QMessageBox.warning(
                self,
                "Chưa cấu hình đủ file nguồn",
                "Vui lòng chọn đủ 5 file Excel nguồn:\n\n"
                + "\n".join(f"• {m}" for m in missing),
            )
            return
        self._save_user_config()
        self._refresh_deploy_summary()
        self.stack.setCurrentIndex(self.PAGE_DEPLOY)
        self._update_step_label()
        if self.thread is None or not self.thread.isRunning():
            self.btn_run.setEnabled(True)

    def _update_step_label(self):
        if self.stack.currentIndex() == self.PAGE_CONFIG:
            self.lbl_step.setText("Bước 1/2 — Cấu hình file nguồn liên kết")
        else:
            self.lbl_step.setText("Bước 2/2 — Triển khai công lương & xem nhật ký")

    def _update_deploy_ui(self):
        """Cập nhật trạng thái trên màn hình cấu hình (nút Tiếp tục, thông báo thiếu file)."""
        complete, missing = self._is_source_config_complete()
        total = len(SOURCE_FILE_DEFINITIONS)
        configured = total - len(missing)

        if self.stack.currentIndex() == self.PAGE_DEPLOY:
            self._refresh_deploy_summary()

        if complete:
            self.lbl_config_status.setVisible(False)
            self.btn_continue.setVisible(True)
        else:
            self.btn_continue.setVisible(False)
            self.lbl_config_status.setVisible(True)
            if configured == 0:
                status_text = (
                    f"⚠️ Vui lòng cấu hình đủ {total} file nguồn phía trên."
                )
            else:
                missing_short = "\n".join(f"  • {m}" for m in missing)
                status_text = (
                    f"⚠️ Đã cấu hình {configured}/{total} file. "
                    f"Còn thiếu ({len(missing)}):\n{missing_short}"
                )
            self.lbl_config_status.setText(status_text)

    def _build_user_config(self) -> dict:
        """Thu thập cấu hình từ các dòng SourceFileRow"""
        result = {}
        for key, row in self.source_rows.items():
            cfg = row.get_config()
            if cfg:
                result[key] = cfg
        return result if result else None
    
    def _save_user_config(self):
        """Lưu cấu hình hiện tại ra file JSON"""
        try:
            config_data = {}
            for key, row in self.source_rows.items():
                cfg = row.get_config()
                if cfg:
                    config_data[key] = cfg
            
            config_path = _get_user_config_path()
            os.makedirs(os.path.dirname(config_path), exist_ok=True)
            with open(config_path, 'w', encoding='utf-8') as f:
                json.dump(config_data, f, ensure_ascii=False, indent=2)
        except Exception as e:
            print(f"⚠️ Không thể lưu user config: {e}")
    
    def _load_user_config(self):
        """Tải cấu hình từ file JSON và khôi phục giao diện"""
        try:
            config_path = _get_user_config_path()
            if not os.path.exists(config_path):
                return
            
            with open(config_path, 'r', encoding='utf-8') as f:
                config_data = json.load(f)
            
            for key, cfg in config_data.items():
                if key in self.source_rows:
                    self.source_rows[key].set_config(cfg)
        except Exception as e:
            print(f"⚠️ Không thể tải user config: {e}")

    def log_callback(self, message: str):
        """Append message to log area"""
        self.log_area.append(message)

    def progress_callback(self, value: int):
        """Update progress bar"""
        self.progress_bar.setValue(value)

    def start_processing(self):
        """
        Bắt đầu triển khai công lương từ metadata.
        Thu thập user_config và lưu trước khi chạy.
        """
        complete, missing = self._is_source_config_complete()
        if not complete:
            QMessageBox.warning(
                self,
                "Chưa cấu hình đủ file nguồn",
                "Vui lòng chọn đủ 5 file Excel nguồn trước khi triển khai:\n\n"
                + "\n".join(f"• {m}" for m in missing),
            )
            return

        month = int(self.combo_month.currentText())
        year = int(self.combo_year.currentText())

        # Thu thập và lưu cấu hình
        user_config = self._build_user_config()
        self._save_user_config()

        if self.stack.currentIndex() != self.PAGE_DEPLOY:
            self._go_to_deploy_page()

        self.btn_run.setEnabled(False)
        self.btn_edit_config.setEnabled(False)
        self.log_area.clear()
        self.progress_bar.setValue(0)
        self.progress_bar.setVisible(True)

        # Khởi tạo thread xử lý
        self.thread = SalaryDeployThread(year, month, user_config)
        self.thread.log_signal.connect(self.log_callback)
        self.thread.progress_signal.connect(self.progress_callback)
        self.thread.finished_signal.connect(self.on_deployment_finished)
        self.thread.start()

    def on_deployment_finished(self, success: bool, message: str):
        """
        Xử lý kết thúc deployment.
        """
        if success:
            QMessageBox.information(
                self, "✅ Thành Công",
                f"Triển khai công lương thành công!\n\n{message}"
            )
        else:
            QMessageBox.critical(
                self, "❌ Lỗi",
                f"Lỗi triển khai:\n\n{message}"
            )
        
        self.progress_bar.setVisible(False)
        self.btn_edit_config.setEnabled(True)
        self._update_deploy_ui()
