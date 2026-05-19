import os
import json
from PySide6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QLabel, QLineEdit,
    QPushButton, QTableWidget, QTableWidgetItem, QHeaderView,
    QMessageBox, QProgressBar, QTextEdit, QFrame, QSizePolicy
)
from PySide6.QtCore import Qt, Slot
from PySide6.QtGui import QFont, QColor
from ui.tab0_template_config.column_config_thread import TemplateUpdateThread
from logic.template_updater import get_column_letter

class ColumnConfigWidget(QWidget):
    """
    Giao diện Quản Lý Cột Tiêu Đề cho Template (Tab 1).
    Cho phép thêm, sửa, xóa, sắp xếp các cột từ F đến AC và đồng bộ vào 2 file template.
    """
    
    DEFAULT_COLUMNS = [
        "Tổ 1", "Tổ 2", "Tổ 3", "Tổ 4", "Tổ 5", "Tổ 6",
        "Tổ 7", "Tổ 8", "Tổ 9", "Tổ 10", "Tổ 11", "Tổ 12",
        "Là TP", "Kiểm hoá", "Cơ động", "Cắt", "Hoàn thiện", "Tái chế"
    ]

    def __init__(self):
        super().__init__()
        self.setObjectName("ColumnConfigWidget")
        
        # Bắt buộc QWidget vẽ nền stylesheet (quan trọng để có nền trắng full màn hình)
        self.setAttribute(Qt.WA_StyledBackground, True)
        
        # Xác định đường dẫn file cấu hình JSON lưu trữ
        current_dir = os.path.abspath(os.path.dirname(__file__))
        self.config_path = os.path.abspath(os.path.join(current_dir, "..", "..", "Template", "sx", "template_columns.json"))
        
        # Tải danh sách cột
        self.columns = []
        self.load_columns_from_json()

        # Áp dụng stylesheet hiện đại, đồng bộ với ứng dụng
        self.setStyleSheet("""
            QWidget#ColumnConfigWidget { background-color: white; border-radius: 8px; }
            QLabel { font-size: 13px; font-weight: bold; color: #374151; }
            QLineEdit { padding: 8px; border: 1px solid #D1D5DB; border-radius: 5px; font-size: 13px; }
            QPushButton { background-color: #3B82F6; color: white; padding: 8px 15px; border-radius: 6px; font-weight: bold; font-size: 13px; border: none; }
            QPushButton:hover { background-color: #2563EB; }
            QPushButton:pressed { background-color: #1D4ED8; }
            QPushButton#DeleteBtn { background-color: #EF4444; }
            QPushButton#DeleteBtn:hover { background-color: #DC2626; }
            QPushButton#ResetBtn { background-color: #F59E0B; }
            QPushButton#ResetBtn:hover { background-color: #D97706; }
            QPushButton#ProcessBtn { background-color: #10B981; font-size: 15px; padding: 12px; border-radius: 8px; }
            QPushButton#ProcessBtn:hover { background-color: #059669; }
            QProgressBar { border: 1px solid #E5E7EB; border-radius: 5px; text-align: center; height: 20px; font-weight: bold; }
            QTextEdit { background-color: #1E1E1E; color: #10B981; font-family: Consolas; padding: 10px; font-size: 12px; }
            QTableWidget { border: 1px solid #E5E7EB; border-radius: 5px; background-color: white; gridline-color: #F3F4F6; }
            QTableWidget::item { padding: 5px; }
            QHeaderView::section { background-color: #F3F4F6; color: #374151; font-weight: bold; border: 1px solid #E5E7EB; }
        """)

        self.init_ui()

    def init_ui(self):
        layout = QVBoxLayout(self)
        layout.setContentsMargins(15, 15, 15, 15)
        layout.setSpacing(10)

        # Tiêu đề Tab
        title = QLabel("🛠️ CẤU HÌNH CỘT TIÊU ĐỀ TEMPLATE")
        title.setFont(QFont("Arial", 16, QFont.Bold))
        title.setAlignment(Qt.AlignCenter)
        layout.addWidget(title)

        # Mô tả hướng dẫn
        desc = QLabel(
            "Mỗi tháng công ty có thể thay đổi danh sách tổ hoặc bộ phận nhận sản lượng.\n"
            "Chức năng này cho phép bạn cấu hình các cột tương ứng từ cột F (Cột số 6) đến cột AC (Cột số 29) trong sheet 'Bang TH nop'.\n"
            "Mọi cột không sử dụng ở phía sau sẽ tự động được ẩn đi trong cả 2 template (to_may_template & kiem_hoa_template)."
        )
        desc.setFont(QFont("Segoe UI", 10))
        desc.setStyleSheet("color: #6B7280; font-weight: normal; margin-bottom: 5px;")
        desc.setAlignment(Qt.AlignCenter)
        layout.addWidget(desc)

        # Layout chính chứa bảng và bảng điều khiển
        main_content_layout = QHBoxLayout()
        main_content_layout.setSpacing(15)

        # 1. Bảng danh sách cột bên trái
        self.table = QTableWidget()
        self.table.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        self.table.setColumnCount(3)
        self.table.setHorizontalHeaderLabels(["STT", "Cột Excel", "Tên Tiêu Đề Cột"])
        self.table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeToContents)
        self.table.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeToContents)
        self.table.horizontalHeader().setSectionResizeMode(2, QHeaderView.Stretch)
        self.table.setSelectionBehavior(QTableWidget.SelectRows)
        self.table.setSelectionMode(QTableWidget.SingleSelection)
        self.table.setFont(QFont("Segoe UI", 10))
        self.table.itemSelectionChanged.connect(self.on_selection_changed)
        main_content_layout.addWidget(self.table, stretch=3)

        # 2. Khung thao tác bên phải
        control_frame = QFrame()
        control_frame.setFrameShape(QFrame.StyledPanel)
        control_frame.setStyleSheet("QFrame { background-color: #F9FAFB; border-radius: 8px; border: 1px solid #E5E7EB; } QLabel { background: transparent; } QPushButton { border-radius: 5px; }")
        control_layout = QVBoxLayout(control_frame)
        control_layout.setContentsMargins(12, 12, 12, 12)
        control_layout.setSpacing(10)

        # Nhập liệu thêm/sửa
        control_layout.addWidget(QLabel("✏️ Nhập Tên Tiêu Đề:"))
        self.txt_title = QLineEdit()
        self.txt_title.setPlaceholderText("Nhập tên tổ/bộ phận (VD: Tổ 13, Là TP...)")
        self.txt_title.returnPressed.connect(self.add_column)
        control_layout.addWidget(self.txt_title)

        # Các nút bấm chức năng
        self.btn_add = QPushButton("➕ Thêm Mới")
        self.btn_add.clicked.connect(self.add_column)
        control_layout.addWidget(self.btn_add)

        self.btn_edit = QPushButton("📝 Sửa Tiêu Đề")
        self.btn_edit.setEnabled(False)
        self.btn_edit.clicked.connect(self.edit_column)
        control_layout.addWidget(self.btn_edit)

        self.btn_delete = QPushButton("❌ Xóa Tiêu Đề")
        self.btn_delete.setObjectName("DeleteBtn")
        self.btn_delete.setEnabled(False)
        self.btn_delete.clicked.connect(self.delete_column)
        control_layout.addWidget(self.btn_delete)

        # Đường kẻ phân cách
        line = QFrame()
        line.setFrameShape(QFrame.HLine)
        line.setFrameShadow(QFrame.Sunken)
        line.setStyleSheet("background-color: #E5E7EB; max-height: 1px;")
        control_layout.addWidget(line)

        # Sắp xếp thứ tự
        control_layout.addWidget(QLabel("↕️ Sắp Xếp Thứ Tự:"))
        self.btn_up = QPushButton("🔼 Di Chuyển Lên")
        self.btn_up.setEnabled(False)
        self.btn_up.clicked.connect(self.move_up)
        control_layout.addWidget(self.btn_up)

        self.btn_down = QPushButton("🔽 Di Chuyển Xuống")
        self.btn_down.setEnabled(False)
        self.btn_down.clicked.connect(self.move_down)
        control_layout.addWidget(self.btn_down)

        control_layout.addStretch()

        # Đặt lại mặc định
        self.btn_reset = QPushButton("🔄 Đặt Lại Mặc Định")
        self.btn_reset.setObjectName("ResetBtn")
        self.btn_reset.clicked.connect(self.reset_to_defaults)
        control_layout.addWidget(self.btn_reset)

        main_content_layout.addWidget(control_frame, stretch=1)
        layout.addLayout(main_content_layout, stretch=4)

        # Nút thực hiện cập nhật templates
        self.btn_run = QPushButton("⚡ CẬP NHẬT CỘT TIÊU ĐỀ VÀO CẢ 2 TEMPLATE")
        self.btn_run.setObjectName("ProcessBtn")
        self.btn_run.clicked.connect(self.start_updating_templates)
        layout.addWidget(self.btn_run)

        # Thanh tiến trình
        self.progress_bar = QProgressBar()
        self.progress_bar.setValue(0)
        layout.addWidget(self.progress_bar)

        # Vùng nhật ký hệ thống
        layout.addWidget(QLabel("📋 Nhật Ký Xử Lý:"))
        self.log_area = QTextEdit()
        self.log_area.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        self.log_area.setReadOnly(True)
        self.log_area.append("ℹ️ Hệ thống đã sẵn sàng. Danh sách cột đã được tải thành công.")
        layout.addWidget(self.log_area, stretch=2)

        # Hiển thị dữ liệu lên bảng
        self.refresh_table()
        
        # Thiết lập Layout rõ ràng cho Widget
        self.setLayout(layout)

    # --- LOGIC XỬ LÝ DỮ LIỆU JSON ---
    
    def load_columns_from_json(self):
        """Tải danh sách cột từ file JSON, nếu chưa có thì dùng mặc định."""
        if os.path.exists(self.config_path):
            try:
                with open(self.config_path, "r", encoding="utf-8") as f:
                    data = json.load(f)
                    self.columns = data.get("columns", self.DEFAULT_COLUMNS.copy())
                    return
            except Exception as e:
                print(f"Lỗi đọc JSON: {e}")
        
        # Nếu chưa có file hoặc lỗi, khởi tạo với danh sách mặc định và lưu luôn
        self.columns = self.DEFAULT_COLUMNS.copy()
        self.save_columns_to_json()

    def save_columns_to_json(self):
        """Lưu danh sách cột và ánh xạ vị trí Excel sang file JSON."""
        # Tạo thư mục nếu chưa tồn tại
        os.makedirs(os.path.dirname(self.config_path), exist_ok=True)
        
        mapping = {}
        for idx, col_name in enumerate(self.columns):
            col_letter = get_column_letter(6 + idx)  # Bắt đầu từ cột F (chỉ số 6)
            mapping[col_name] = col_letter

        data = {
            "columns": self.columns,
            "mapping": mapping
        }

        try:
            with open(self.config_path, "w", encoding="utf-8") as f:
                json.dump(data, f, ensure_ascii=False, indent=2)
        except Exception as e:
            QMessageBox.critical(self, "Lỗi Lưu Trữ", f"Không thể ghi file cấu hình JSON: {e}")

    # --- ĐIỀU KHIỂN BẢNG VÀ GIAO DIỆN ---
    
    def refresh_table(self):
        """Cập nhật dữ liệu từ list `self.columns` lên QTableWidget."""
        self.table.setRowCount(0)
        for idx, col_name in enumerate(self.columns):
            row_idx = self.table.rowCount()
            self.table.insertRow(row_idx)

            # 1. Số thứ tự (STT)
            stt_item = QTableWidgetItem(str(row_idx + 1))
            stt_item.setTextAlignment(Qt.AlignCenter)
            stt_item.setFlags(stt_item.flags() ^ Qt.ItemIsEditable)
            self.table.setItem(row_idx, 0, stt_item)

            # 2. Cột Excel tương ứng (F, G, H...)
            col_letter = get_column_letter(6 + idx)
            letter_item = QTableWidgetItem(col_letter)
            letter_item.setTextAlignment(Qt.AlignCenter)
            letter_item.setFlags(letter_item.flags() ^ Qt.ItemIsEditable)
            letter_item.setForeground(QColor("#2563EB"))  # Highlight màu xanh dương sang trọng
            self.table.setItem(row_idx, 1, letter_item)

            # 3. Tên tiêu đề cột
            title_item = QTableWidgetItem(col_name)
            title_item.setFlags(title_item.flags() ^ Qt.ItemIsEditable)
            self.table.setItem(row_idx, 2, title_item)

        # Reset các trường nhập liệu và nút
        self.txt_title.clear()
        self.btn_edit.setEnabled(False)
        self.btn_delete.setEnabled(False)
        self.btn_up.setEnabled(False)
        self.btn_down.setEnabled(False)

    def on_selection_changed(self):
        """Xử lý sự kiện khi click chọn dòng trên bảng."""
        selected_rows = self.table.selectionModel().selectedRows()
        if not selected_rows:
            self.btn_edit.setEnabled(False)
            self.btn_delete.setEnabled(False)
            self.btn_up.setEnabled(False)
            self.btn_down.setEnabled(False)
            return

        row = selected_rows[0].row()
        col_name = self.columns[row]
        self.txt_title.setText(col_name)
        
        self.btn_edit.setEnabled(True)
        self.btn_delete.setEnabled(True)
        self.btn_up.setEnabled(row > 0)
        self.btn_down.setEnabled(row < len(self.columns) - 1)

    # --- CÁC HÀM THAO TÁC CỘT (THÊM, SỬA, XÓA, SẮP XẾP) ---
    
    @Slot()
    def add_column(self):
        """Thêm một cột mới vào danh sách."""
        text = self.txt_title.text().strip()
        if not text:
            QMessageBox.warning(self, "Cảnh Báo", "Vui lòng nhập tên tiêu đề cột trước khi thêm.")
            return

        # Giới hạn tối đa 24 cột (từ F đến AC)
        if len(self.columns) >= 24:
            QMessageBox.warning(
                self, "Giới Hạn Cột", 
                "Excel template chỉ hỗ trợ tối đa 24 cột tiêu đề sản lượng (từ F đến AC).\n"
                "Bạn không thể thêm nhiều hơn 24 cột."
            )
            return

        if text in self.columns:
            QMessageBox.warning(self, "Trùng Tiêu Đề", f"Tiêu đề '{text}' đã tồn tại trong danh sách.")
            return

        self.columns.append(text)
        self.save_columns_to_json()
        self.refresh_table()
        self.log_area.append(f"➕ Đã thêm tiêu đề cột mới: '{text}' (Ánh xạ vào cột {get_column_letter(6 + len(self.columns) - 1)})")
        
        # Cuộn bảng xuống dòng cuối cùng
        self.table.scrollToBottom()

    @Slot()
    def edit_column(self):
        """Sửa tên tiêu đề cột đang được chọn."""
        selected_rows = self.table.selectionModel().selectedRows()
        if not selected_rows:
            return

        row = selected_rows[0].row()
        old_text = self.columns[row]
        new_text = self.txt_title.text().strip()

        if not new_text:
            QMessageBox.warning(self, "Cảnh Báo", "Tên tiêu đề cột không được bỏ trống.")
            return

        if new_text == old_text:
            return

        if new_text in self.columns and self.columns.index(new_text) != row:
            QMessageBox.warning(self, "Trùng Tiêu Đề", f"Tiêu đề '{new_text}' đã được sử dụng ở dòng khác.")
            return

        self.columns[row] = new_text
        self.save_columns_to_json()
        self.refresh_table()
        self.log_area.append(f"✏️ Đã đổi tiêu đề dòng {row + 1} từ '{old_text}' thành '{new_text}'")

    @Slot()
    def delete_column(self):
        """Xóa cột đang chọn ra khỏi danh sách."""
        selected_rows = self.table.selectionModel().selectedRows()
        if not selected_rows:
            return

        row = selected_rows[0].row()
        col_name = self.columns[row]

        reply = QMessageBox.question(
            self, "Xác Nhận Xóa",
            f"Bạn có chắc chắn muốn xóa cột '{col_name}' này không?\n"
            "Điều này sẽ dịch chuyển tất cả các cột phía sau lên 1 vị trí cột Excel.",
            QMessageBox.Yes | QMessageBox.No, QMessageBox.No
        )

        if reply == QMessageBox.Yes:
            deleted_name = self.columns.pop(row)
            self.save_columns_to_json()
            self.refresh_table()
            self.log_area.append(f"❌ Đã xóa tiêu đề cột: '{deleted_name}'")

    @Slot()
    def move_up(self):
        """Di chuyển cột được chọn lên trên 1 dòng."""
        selected_rows = self.table.selectionModel().selectedRows()
        if not selected_rows:
            return

        row = selected_rows[0].row()
        if row == 0:
            return

        # Tráo đổi vị trí
        self.columns[row], self.columns[row - 1] = self.columns[row - 1], self.columns[row]
        self.save_columns_to_json()
        self.refresh_table()
        
        # Chọn lại dòng đã di chuyển
        self.table.selectRow(row - 1)
        self.log_area.append(f"🔼 Di chuyển '{self.columns[row - 1]}' lên dòng {row}")

    @Slot()
    def move_down(self):
        """Di chuyển cột được chọn xuống dưới 1 dòng."""
        selected_rows = self.table.selectionModel().selectedRows()
        if not selected_rows:
            return

        row = selected_rows[0].row()
        if row == len(self.columns) - 1:
            return

        # Tráo đổi vị trí
        self.columns[row], self.columns[row + 1] = self.columns[row + 1], self.columns[row]
        self.save_columns_to_json()
        self.refresh_table()
        
        # Chọn lại dòng đã di chuyển
        self.table.selectRow(row + 1)
        self.log_area.append(f"🔽 Di chuyển '{self.columns[row + 1]}' xuống dòng {row + 2}")

    @Slot()
    def reset_to_defaults(self):
        """Khôi phục danh sách cột tiêu đề về danh sách mặc định ban đầu."""
        reply = QMessageBox.question(
            self, "Đặt Lại Mặc Định",
            "Bạn có chắc chắn muốn đặt lại danh sách cột tiêu đề về mặc định của nhà máy không?\n"
            "Toàn bộ cấu hình tùy chỉnh hiện tại sẽ bị xóa sạch.",
            QMessageBox.Yes | QMessageBox.No, QMessageBox.No
        )

        if reply == QMessageBox.Yes:
            self.columns = self.DEFAULT_COLUMNS.copy()
            self.save_columns_to_json()
            self.refresh_table()
            self.log_area.append("🔄 Đã khôi phục danh sách cột tiêu đề về mặc định thành công.")

    # --- HÀM THỰC THI CHẠY NGẦM CẬP NHẬT TEMPLATES ---
    
    @Slot()
    def start_updating_templates(self):
        """Bắt đầu chạy luồng cập nhật tiêu đề cột vào 2 template excel."""
        if not self.columns:
            QMessageBox.warning(self, "Danh Sách Trống", "Danh sách tiêu đề cột trống. Vui lòng thêm ít nhất 1 cột.")
            return

        reply = QMessageBox.question(
            self, "Xác Nhận Xử Lý",
            "Hệ thống sẽ thực hiện cập nhật đồng bộ các cột tiêu đề mới vào cả 2 template:\n"
            "- Template\\sx\\to_may_template.xlsx\n"
            "- Template\\sx\\kiem_hoa_template.xlsx\n\n"
            "LƯU Ý: Vui lòng đảm bảo các file Excel template này đang được đóng.\n"
            "Bạn có muốn tiếp tục?",
            QMessageBox.Yes | QMessageBox.No, QMessageBox.No
        )

        if reply != QMessageBox.Yes:
            return

        self.log_area.clear()
        self.log_area.append("⏳ Đang chuẩn bị cập nhật Template Excel...")
        self.progress_bar.setValue(0)
        self.set_controls_enabled(False)

        # Khởi tạo và chạy luồng ngầm
        self.thread = TemplateUpdateThread(self.columns)
        self.thread.log_signal.connect(self.log_area.append)
        self.thread.progress_signal.connect(self.progress_bar.setValue)
        self.thread.finished_signal.connect(self.on_update_finished)
        self.thread.start()

    def set_controls_enabled(self, enabled):
        """Khóa/Mở khóa giao diện khi đang chạy ngầm."""
        self.txt_title.setEnabled(enabled)
        self.btn_add.setEnabled(enabled)
        self.btn_reset.setEnabled(enabled)
        self.btn_run.setEnabled(enabled)
        self.table.setEnabled(enabled)
        if enabled:
            self.on_selection_changed()
        else:
            self.btn_edit.setEnabled(False)
            self.btn_delete.setEnabled(False)
            self.btn_up.setEnabled(False)
            self.btn_down.setEnabled(False)

    def on_update_finished(self, success, message):
        """Nhận kết quả từ luồng chạy ngầm khi hoàn thành."""
        self.set_controls_enabled(True)
        if success:
            QMessageBox.information(
                self, "Thành Công",
                "Đã cập nhật đồng bộ cột tiêu đề thành công vào 2 file template!\n"
                "Kể từ bây giờ, các tab phía sau sẽ tự động sử dụng cấu hình cột mới này."
            )
            self.progress_bar.setValue(100)
        else:
            QMessageBox.critical(
                self, "Lỗi Cập Nhật",
                f"Đã xảy ra lỗi khi cập nhật template:\n{message}"
            )
            self.progress_bar.setValue(0)
