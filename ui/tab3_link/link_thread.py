from PySide6.QtCore import QThread, Signal
from logic.excel_integration import process_excel_integration

class IntegrationThread(QThread):
    """
    Luồng xử lý chạy ngầm để tích hợp dữ liệu JSON vào Excel hàng loạt.
    """
    log_signal = Signal(str)
    progress_signal = Signal(int)
    finished_signal = Signal(bool, str)

    def __init__(self, json_files, excel_dir, password):
        # Khởi tạo luồng tích hợp
        super().__init__()
        self.json_files = json_files
        self.excel_dir = excel_dir
        self.password = password

    def log_callback(self, message):
        # Gửi log về UI
        self.log_signal.emit(message)
        
    def progress_callback(self, value):
        # Gửi tiến độ về UI
        self.progress_signal.emit(value)

    def run(self):
        # Thực hiện gọi logic tích hợp Excel
        success, message = process_excel_integration(
            self.json_files,
            self.excel_dir,
            self.password,
            self.log_callback,
            self.progress_callback
        )
        self.finished_signal.emit(success, message)
