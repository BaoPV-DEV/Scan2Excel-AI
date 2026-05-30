from PySide6.QtCore import QThread, Signal
from app.logic.excel_integration import process_excel_integration


class IntegrationThread(QThread):
    """
    Thread wrapper cho quá trình integrate JSON -> Excel.
    (Optimized: reduced indirection, cleaner callbacks, safer threading)
    """

    log_signal = Signal(str)
    progress_signal = Signal(int)
    finished_signal = Signal(bool, str)

    def __init__(self, json_root, excel_dir, password, single_dept=None, parent=None):
        super().__init__(parent)

        self.json_root = json_root
        self.excel_dir = excel_dir
        self.password = password
        self.single_dept = single_dept

    # =========================
    # CALLBACK WRAPPERS (inline optimized)
    # =========================
    def _log(self, message: str):
        self.log_signal.emit(message)

    def _progress(self, value: int):
        self.progress_signal.emit(value)

    # =========================
    # THREAD ENTRY
    # =========================
    def run(self):
        try:
            print(f"Starting integration: json_root={self.json_root}, excel_dir={self.excel_dir}, single_dept={self.single_dept}")
            success, message = process_excel_integration(
                self.json_root,
                self.excel_dir,
                self._log,
                self._progress,
                self.password,
                self.single_dept
            )
        except Exception as e:
            # đảm bảo thread không chết silent
            success = False
            message = str(e)
            self._log(f"❌ Integration error: {message}")

        self.finished_signal.emit(success, message)