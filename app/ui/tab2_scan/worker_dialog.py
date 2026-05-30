from PySide6.QtWidgets import QDialog, QVBoxLayout, QHBoxLayout, QPushButton, QLabel, QTableWidget, QTableWidgetItem, QHeaderView, QDialogButtonBox
from PySide6.QtCore import Qt

# Cho phép người dùng chỉnh sửa danh sách nhân viên và số lượng sản phẩm của từng người trong một công đoạn.
class WorkerEditorDialog(QDialog):
    # Khởi tạo hộp thoại chỉnh sửa nhân viên
    def __init__(self, workers, title_info="", parent=None):
        super().__init__(parent)
        self.setWindowTitle("CHI TIẾT NGƯỜI THỰC HIỆN")
        self.resize(550, 450)
        self.workers = workers.copy()
        self.title_info = title_info
        self.init_ui()

    # Thiết lập các thành phần giao diện cho hộp thoại
    def init_ui(self):
        layout = QVBoxLayout(self)
        
        info_lbl = QLabel(f"📍 Đang sửa nhân sự cho: {self.title_info}")
        info_lbl.setStyleSheet("font-weight: bold; color: #2563EB; font-size: 14px; padding: 5px;")
        layout.addWidget(info_lbl)
        
        self.table = QTableWidget()
        self.table.setColumnCount(3)
        self.table.setHorizontalHeaderLabels(["Nhân Viên", "Mã NV", "Số Lượng"])
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        
        self.load_data()
        layout.addWidget(self.table)
        
        btn_layout = QHBoxLayout()
        add_btn = QPushButton("➕ Thêm người")
        add_btn.clicked.connect(self.add_row)
        remove_btn = QPushButton("➖ Xóa người")
        remove_btn.clicked.connect(self.remove_row)
        
        btn_layout.addWidget(add_btn)
        btn_layout.addWidget(remove_btn)
        btn_layout.addStretch()
        layout.addLayout(btn_layout)
        
        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        buttons.accepted.connect(self.save_and_accept)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)

    # Đổ dữ liệu từ danh sách vào bảng hiển thị
    def load_data(self):
        self.table.setRowCount(len(self.workers))
        for i, w in enumerate(self.workers):
            self.table.setItem(i, 0, QTableWidgetItem(str(w.get("nhan_vien", ""))))
            self.table.setItem(i, 1, QTableWidgetItem(str(w.get("ma_nhan_vien", ""))))
            self.table.setItem(i, 2, QTableWidgetItem(str(w.get("so_luong", ""))))

    # Thêm một hàng dữ liệu mới vào bảng
    def add_row(self):
        row = self.table.rowCount()
        self.table.insertRow(row)
        self.table.setItem(row, 0, QTableWidgetItem(""))
        self.table.setItem(row, 1, QTableWidgetItem(""))
        self.table.setItem(row, 2, QTableWidgetItem("0"))

    # Xóa hàng dữ liệu đang được chọn
    def remove_row(self):
        curr = self.table.currentRow()
        if curr >= 0:
            self.table.removeRow(curr)

    # Lưu dữ liệu từ bảng vào bộ nhớ và đóng hộp thoại
    def save_and_accept(self):
        new_workers = []
        for r in range(self.table.rowCount()):
            name = self.table.item(r, 0).text().strip()
            ma = self.table.item(r, 1).text().strip()
            sl = self.table.item(r, 2).text().strip()
            if name or ma:
                try: sl_int = int(sl)
                except: sl_int = 0
                new_workers.append({
                    "nhan_vien": name,
                    "ma_nhan_vien": ma,
                    "so_luong": sl_int
                })
        self.workers = new_workers
        self.accept()

    # Trả về danh sách nhân viên sau khi đã chỉnh sửa
    def get_workers(self):
        return self.workers
