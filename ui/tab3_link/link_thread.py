from PySide6.QtCore import QThread, Signal
from logic.excel_integration import process_excel_integration

# Luồng xử lý chạy ngầm để tích hợp dữ liệu JSON vào Excel hàng loạt.
class IntegrationThread(QThread):
    log_signal = Signal(str)
    progress_signal = Signal(int)
    finished_signal = Signal(bool, str)

    # Khởi tạo luồng tích hợp
    def __init__(self, json_root, excel_dir, password, single_dept=None):
        super().__init__()
        self.json_root = json_root
        self.excel_dir = excel_dir
        self.password = password
        self.single_dept = single_dept

    # Gửi log về UI
    def log_callback(self, message):
        self.log_signal.emit(message)
        
    # Gửi tiến độ về UI
    def progress_callback(self, value):
        self.progress_signal.emit(value)

    # Thực hiện gọi logic tích hợp Excel
    def run(self):
        success, message = process_excel_integration(
            self.json_root,
            self.excel_dir,
            self.log_callback,
            self.progress_callback,
            self.password,
            self.single_dept
        )
        self.finished_signal.emit(success, message)
