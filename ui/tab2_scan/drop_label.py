from PySide6.QtWidgets import QLabel
from PySide6.QtCore import Qt

# Lớp hỗ trợ hiển thị vùng kéo thả file ảnh.
class DropLabel(QLabel):
    # Khởi tạo vùng kéo thả với style đứt đoạn
    def __init__(self, text, parent=None):
        super().__init__(text, parent)
        self.setAlignment(Qt.AlignCenter)
        self.setStyleSheet("""
            QLabel {
                border: 2px dashed #D1D5DB;
                border-radius: 10px;
                color: #6B7280;
                background-color: #F9FAFB;
                font-size: 14px;
            }
        """)
        self.setMinimumHeight(100)
