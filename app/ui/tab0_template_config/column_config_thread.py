from PySide6.QtCore import QThread, Signal, Slot
from app.logic.template_updater import update_templates_with_headers


class TemplateUpdateThread(QThread):
    log_signal = Signal(str)
    progress_signal = Signal(int)
    finished_signal = Signal(bool, str)

    def __init__(self, titles):
        super().__init__()
        self.titles = titles
        self._is_running = True

    # =========================
    # PUBLIC CONTROL
    # =========================
    def stop(self):
        """Cho phép UI request stop thread an toàn."""
        self._is_running = False

    # =========================
    # CALLBACK WRAPPERS
    # =========================
    def log_callback(self, message: str):
        if self._is_running:
            self.log_signal.emit(message)

    def progress_callback(self, value: int):
        if self._is_running:
            self.progress_signal.emit(value)

    # =========================
    # THREAD ENTRY POINT
    # =========================
    def run(self):
        try:
            if not self._is_running:
                self.finished_signal.emit(False, "Cancelled before start")
                return

            success, message = update_templates_with_headers(
                self.titles,
                self.log_callback,
                self.progress_callback
            )

            if self._is_running:
                self.finished_signal.emit(success, message)

        except Exception as e:
            self.finished_signal.emit(False, str(e))