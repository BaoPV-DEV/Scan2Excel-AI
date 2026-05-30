import sys
import ctypes

from PySide6.QtWidgets import QApplication, QStyleFactory
from PySide6.QtGui import QIcon

from app.utils.logger import setup_logger
from app.utils.paths import resource_path
from app.ui.gui import MainWindow


def main():
    # =========================
    # LOGGING FIRST
    # =========================
    setup_logger()

    # =========================
    # FIX TASKBAR ICON WINDOWS
    # =========================
    myappid = 'sd.scan2excel.ai.1'
    try:
        ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(myappid)
    except Exception:
        pass

    # =========================
    # QT APP INIT
    # =========================
    app = QApplication(sys.argv)

    # 🔥 FIX QUAN TRỌNG: ép Fusion để QSS luôn hoạt động ổn định
    app.setStyle(QStyleFactory.create("Fusion"))

    # =========================
    # GLOBAL ICON
    # =========================
    icon = QIcon(resource_path("resources/icons/app_icon.png"))
    app.setWindowIcon(icon)

    # =========================
    # OPTIONAL GLOBAL STYLE (ổn định hơn per-widget)
    # =========================
    app.setStyleSheet("""
        QWidget {
            font-family: Segoe UI;
        }

        QPushButton {
            border-radius: 6px;
        }

        QTabWidget::pane {
            border: 1px solid #D1D5DB;
            background: white;
        }

        QTabBar::tab {
            background: #E5E7EB;
            padding: 10px;
        }

        QTabBar::tab:selected {
            background: white;
            color: #2563EB;
        }
    """)

    # =========================
    # MAIN WINDOW
    # =========================
    window = MainWindow()
    window.setWindowIcon(icon)

    window.showMaximized()

    sys.exit(app.exec())


if __name__ == "__main__":
    main()