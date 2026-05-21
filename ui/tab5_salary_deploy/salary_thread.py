import os
import re
from PySide6.QtCore import QThread, Signal
from logic.salary_config import SALARY_TEMPLATE_FORMULA_CONFIG

class SalaryDeployThread(QThread):
    """
    Luồng xử lý (Thread) triển khai công thức lương link dữ liệu ngoài.
    Chạy ngầm để không làm treo giao diện.
    """
    log_signal = Signal(str)
    progress_signal = Signal(int)
    finished_signal = Signal(bool, str)

    def __init__(self, excel_files_dir, template_key, month, year):
        super().__init__()
        self.excel_files_dir = excel_files_dir
        self.template_key = template_key
        self.month = month
        self.year = year

    def log_callback(self, message):
        """Gửi thông báo nhật ký về giao diện"""
        self.log_signal.emit(message)
    
    def progress_callback(self, value):
        """Cập nhật tiến độ xử lý cho giao diện"""
        self.progress_signal.emit(value)

    def run(self):
        """Chạy quy trình triển khai công lương từ JSON config"""
        try:
            from logic.salary_deployer import apply_salary_formulas_from_json
            
            # Áp dụng công thức lương cho tất cả file Excel trong thư mục
            success, message = apply_salary_formulas_from_json(
                self.excel_files_dir,
                self.template_key,
                self.month,
                self.year,
                self.log_callback,
                self.progress_callback
            )
            self.finished_signal.emit(success, message)
        except Exception as e:
            self.log_callback(f"❌ Lỗi triển khai công lương: {str(e)}")
            import traceback
            self.log_callback(f"{traceback.format_exc()}")
            self.finished_signal.emit(False, f"Lỗi: {str(e)}")
