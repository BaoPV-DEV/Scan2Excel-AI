import os
from PySide6.QtCore import QThread, Signal
from logic.excel_processor import process_excel

class ExcelProcessThread(QThread):
    """
    Luồng xử lý (Thread) chạy ngầm để tách file Excel mà không làm treo giao diện.
    """
    log_signal = Signal(str)
    progress_signal = Signal(int)
    finished_signal = Signal(bool, str)

    # Khởi tạo luồng xử lý
    def __init__(self, input_path, template_path, output_dir, month, year, organize_folders):
        super().__init__()
        self.input_path = input_path
        self.template_path = template_path
        self.output_dir = output_dir
        self.month = month
        self.year = year
        self.organize_folders = organize_folders

    # Gửi thông báo nhật ký về giao diện
    def log_callback(self, message):
        self.log_signal.emit(message)
    
    # Cập nhật tiến độ xử lý cho giao diện
    def progress_callback(self, value):
        self.progress_signal.emit(value)

    # Tự động tạo thư mục theo Năm/Tháng nếu được chọn
    def run(self):
        final_output = self.output_dir
        if self.organize_folders:
            final_output = os.path.join(self.output_dir, self.year, self.month)
            
            # Làm mới thư mục (xóa cũ tạo mới) để đảm bảo dữ liệu sạch
            if os.path.exists(final_output):
                try:
                    import shutil
                    shutil.rmtree(final_output)
                    self.log_callback(f"🧹 Đã xóa thư mục cũ để làm mới dữ liệu: {final_output}")
                except Exception as e:
                    self.log_callback(f"⚠️ Cảnh báo: Không thể xóa thư mục cũ (có thể file đang mở): {str(e)}")

            os.makedirs(final_output, exist_ok=True)
            self.log_callback(f"📂 Đã tạo thư mục lưu trữ: {final_output}")

        # Gọi hàm xử lý logic từ module excel_processor
        success, message = process_excel(
            self.input_path, 
            self.template_path, 
            final_output, 
            self.month, 
            self.year, 
            self.log_callback,
            self.progress_callback
        )
        self.finished_signal.emit(success, message)
