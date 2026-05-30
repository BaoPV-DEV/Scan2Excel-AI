from PySide6.QtCore import QThread, Signal
from app.logic.salary_deployer import deploy_salary_from_metadata
import traceback


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

    def __init__(self, year: int, month: int, user_config: dict | None = None):
        super().__init__()

        self.year = year
        self.month = month
        self.user_config = dict(user_config) if user_config else {}

        self._running = True  # future-proof guard

    # ---------------- callbacks ----------------
    def log_callback(self, message: str):
        self.log_signal.emit(message)

    def progress_callback(self, value: int):
        self.progress_signal.emit(value)

    # ---------------- core ----------------
    def run(self):
        """
        Execute salary deployment pipeline safely.
        """
        try:
            if not self._running:
                return

            success, message = deploy_salary_from_metadata(
                self.year,
                self.month,
                self.log_callback,
                self.progress_callback,
                self.user_config
            )

            self.finished_signal.emit(success, message)

        except Exception as e:
            self.log_callback(f"❌ Lỗi triển khai công lương: {e}")
            self.log_callback(traceback.format_exc())
            self.finished_signal.emit(False, str(e))

    # ---------------- safety ----------------
    def stop(self):
        """Cho phép cancel thread từ UI (future use)"""
        self._running = False