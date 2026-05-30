from PySide6.QtWidgets import QMainWindow, QTabWidget, QWidget, QVBoxLayout, QApplication
from PySide6.QtGui import QIcon, QFont
from app.utils.paths import resource_path


class MainWindow(QMainWindow):

    def __init__(self):
        super().__init__()

        self.setWindowTitle("Idol Linh tính lương")
        self.setGeometry(100, 100, 1000, 750)
        self.setWindowIcon(QIcon(resource_path("resources/icons/app_icon.png")))

        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        central_widget.setStyleSheet("background-color: #F3F4F6;")

        layout = QVBoxLayout(central_widget)
        layout.setContentsMargins(15, 15, 15, 15)

        self.tabs = QTabWidget()
        self.tabs.setFont(QFont("Segoe UI", 11, QFont.Bold))
        layout.addWidget(self.tabs)

        # =========================
        # PLACEHOLDER TAB (lazy)
        # =========================
        self.tabs.addTab(QWidget(), "1. Cập Nhật Mẫu")
        self.tabs.addTab(QWidget(), "2. Tách File Nhân Viên")
        self.tabs.addTab(QWidget(), "3. Quét Ảnh Sản Lượng")
        self.tabs.addTab(QWidget(), "4. Tích Hợp Sản Lượng")
        self.tabs.addTab(QWidget(), "5. Triển Khai Lương")

        self.tab_instances = {}

        self.tabs.currentChanged.connect(self._on_tab_changed)
        self._on_tab_changed(0)  # Load tab đầu tiên ngay khi khởi động

        self.tabs.setStyleSheet("""
            QTabWidget::pane {
                border: 1px solid #D1D5DB;
                background-color: white;
                border-radius: 8px;
            }
            QTabBar::tab {
                background-color: #E5E7EB;
                color: #4B5563;
                padding: 12px 25px;
                margin-right: 5px;
                border-top-left-radius: 8px;
                border-top-right-radius: 8px;
            }
            QTabBar::tab:selected {
                background-color: white;
                color: #2563EB;
                border-top: 3px solid #2563EB;
            }
        """)

    # =========================
    # LAZY LOADING ENGINE
    # =========================
    def _on_tab_changed(self, index: int):
        if index in self.tab_instances:
            return

        if index == 0:
            from app.ui.tab0_template_config.column_config_widget import ColumnConfigWidget
            widget = ColumnConfigWidget()

        elif index == 1:
            from app.ui.tab1_split.split_widget import SplitExcelWidget
            widget = SplitExcelWidget()

        elif index == 2:
            from app.ui.tab2_scan.scan_widget import ScanProductionWidget
            widget = ScanProductionWidget()

        elif index == 3:
            from app.ui.tab3_link.link_widget import LinkExcelWidget
            widget = LinkExcelWidget()

        elif index == 4:
            from app.ui.tab5_salary_deploy.salary_widget import SalaryDeployWidget
            widget = SalaryDeployWidget()

        else:
            return

        self.tab_instances[index] = widget
        self.tabs.removeTab(index)
        self.tabs.insertTab(index, widget, self._tab_name(index))
        self.tabs.setCurrentIndex(index)

    def _tab_name(self, i):
        return [
            "1. Cập Nhật Mẫu",
            "2. Tách File Nhân Viên",
            "3. Quét Ảnh Sản Lượng",
            "4. Tích Hợp Sản Lượng",
            "5. Triển Khai Lương",
        ][i]