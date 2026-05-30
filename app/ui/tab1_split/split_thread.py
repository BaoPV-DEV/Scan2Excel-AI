import os
import traceback
from PySide6.QtCore import QThread, Signal
from app.logic.excel_processor import process_excel


class ExcelProcessThread(QThread):
    """
    Luồng xử lý Excel chạy nền - tối ưu cho production / build EXE
    """
    log_signal = Signal(str)
    progress_signal = Signal(int)
    finished_signal = Signal(bool, str)

    def __init__(self, input_path, output_dir, month, year):
        super().__init__()
        self.input_path = input_path
        self.output_dir = output_dir
        self.month = month
        self.year = year

        self._is_running = True  # safeguard

    def log_callback(self, message: str):
        if self._is_running:
            self.log_signal.emit(str(message))

    def progress_callback(self, value: int):
        if not self._is_running:
            return

        # clamp tránh UI lỗi
        value = max(0, min(100, int(value)))
        self.progress_signal.emit(value)

    def run(self):
        try:
            os.makedirs(self.output_dir, exist_ok=True)
            self.log_callback(f"📂 Thư mục lưu trữ: {self.output_dir}")

            success, message = process_excel(
                self.input_path,
                None,
                self.output_dir,
                self.month,
                self.year,
                self.log_callback,
                self.progress_callback
            )

            self.finished_signal.emit(bool(success), str(message))

        except Exception as e:
            # đảm bảo KHÔNG bao giờ chết thread im lặng
            err = f"{e}\n{traceback.format_exc()}"
            self.log_callback(f"❌ Lỗi thread: {e}")
            self.finished_signal.emit(False, err)

        finally:
            self._is_running = False