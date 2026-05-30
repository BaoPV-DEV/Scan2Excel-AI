from PySide6.QtWidgets import QLabel
from PySide6.QtCore import Qt


class DropLabel(QLabel):
    """
    Label hỗ trợ vùng kéo thả file ảnh (drag & drop UI placeholder).
    """

    _STYLE = """
        QLabel {
            border: 2px dashed #D1D5DB;
            border-radius: 10px;
            color: #6B7280;
            background-color: #F9FAFB;
            font-size: 14px;
        }
    """

    def __init__(self, text, parent=None):
        super().__init__(text, parent)

        # Qt optimization: ổn định background rendering (quan trọng khi embed UI lớn)
        self.setAttribute(Qt.WA_StyledBackground, True)

        self.setAlignment(Qt.AlignCenter)
        self.setMinimumHeight(100)

        # cache style để tránh parse lại string mỗi instance
        self.setStyleSheet(self._STYLE)