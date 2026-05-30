from PySide6.QtCore import QThread, Signal
from app.logic.template_updater import update_templates_with_headers

class TemplateUpdateThread(QThread):
    log_signal = Signal(str)
    progress_signal = Signal(int)
    finished_signal = Signal(bool, str)

    def __init__(self, titles):
        super().__init__()
        self.titles = titles

    def log_callback(self, message):
        self.log_signal.emit(message)

    def progress_callback(self, value):
        self.progress_signal.emit(value)

    def run(self):
        success, message = update_templates_with_headers(
            self.titles,
            self.log_callback,
            self.progress_callback
        )
        self.finished_signal.emit(success, message)
