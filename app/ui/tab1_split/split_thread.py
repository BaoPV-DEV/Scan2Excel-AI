import os
from PySide6.QtCore import QThread, Signal
from app.logic.excel_processor import process_excel

class ExcelProcessThread(QThread):
    """
    Luồng xử lý (Thread) chạy ngầm để tách file Excel mà không làm treo giao diện.
    """
    log_signal = Signal(str)
    progress_signal = Signal(int)
    finished_signal = Signal(bool, str)

    # Khởi tạo luồng xử lý
    def __init__(self, input_path, output_dir, month, year):
        super().__init__()
        self.input_path = input_path
        self.output_dir = output_dir
        self.month = month
        self.year = year

    # Gửi thông báo nhật ký về giao diện
    def log_callback(self, message):
        self.log_signal.emit(message)
    
    # Cập nhật tiến độ xử lý cho giao diện
    def progress_callback(self, value):
        self.progress_signal.emit(value)

    # Chạy quy trình xử lý chính
    def run(self):
        # Đảm bảo thư mục output tồn tại
        os.makedirs(self.output_dir, exist_ok=True)
        self.log_callback(f"📂 Thư mục lưu trữ: {self.output_dir}")

        # Gọi hàm xử lý logic từ module excel_processor
        success, message = process_excel(
            self.input_path, 
            None,  # template_path không còn cần (tự chọn theo phân loại)
            self.output_dir, 
            self.month, 
            self.year, 
            self.log_callback,
            self.progress_callback
        )
        self.finished_signal.emit(success, message)
