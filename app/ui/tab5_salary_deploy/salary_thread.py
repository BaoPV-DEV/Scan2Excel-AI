import os
from PySide6.QtCore import QThread, Signal
from app.logic.salary_deployer import deploy_salary_from_metadata


class SalaryDeployThread(QThread):
    """
    Luồng xử lý (Thread) triển khai công lương từ metadata.
    
    Workflow:
    1. Đọc metadata.json từ Tab 2 output
    2. Xác định template cho từng file
    3. Tuần tự xử lý: fill công thức + linking dữ liệu
    4. Lưu file
    """
    log_signal = Signal(str)
    progress_signal = Signal(int)
    finished_signal = Signal(bool, str)

    def __init__(self, year: int, month: int, user_config: dict = None):
        super().__init__()
        self.year = year
        self.month = month
        self.user_config = user_config

    def log_callback(self, message: str):
        """Gửi thông báo nhật ký về giao diện"""
        self.log_signal.emit(message)
    
    def progress_callback(self, value: int):
        """Cập nhật tiến độ xử lý cho giao diện"""
        self.progress_signal.emit(value)

    def run(self):
        """
        Chạy quy trình triển khai công lương từ metadata.
        Gọi hàm deploy_salary_from_metadata() với user_config.
        """
        try:
            # Gọi hàm triển khai chính
            success, message = deploy_salary_from_metadata(
                self.year,
                self.month,
                self.log_callback,
                self.progress_callback,
                self.user_config
            )
            
            self.finished_signal.emit(success, message)
        
        except Exception as e:
            self.log_callback(f"❌ Lỗi triển khai công lương: {str(e)}")
            import traceback
            self.log_callback(f"{traceback.format_exc()}")
            self.finished_signal.emit(False, f"Lỗi: {str(e)}")
