import sys
from PySide6.QtWidgets import QMainWindow, QTabWidget, QWidget, QVBoxLayout, QApplication
from PySide6.QtGui import QIcon, QFont


# Import các Widget từ các file giao diện thành phần (Cấu trúc mới theo từng Tab)
from ui.tab1_split.split_widget import SplitExcelWidget
from ui.tab2_scan.scan_widget import ScanProductionWidget
from ui.tab3_link.link_widget import LinkExcelWidget

class MainWindow(QMainWindow):
    """
    Lớp cửa sổ chính của ứng dụng.
    Quản lý các tab chức năng: Tách file, Scan ảnh, và Tích hợp Excel.
    """
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Idol Linh tính lương")
        self.setGeometry(100, 100, 1000, 750)
        self.setWindowIcon(QIcon("app_icon.png"))
        self.showMaximized() # Mở to toàn màn hình
        
        # Widget trung tâm và Layout chính
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        central_widget.setStyleSheet("background-color: #F3F4F6;") # Màu nền xám nhạt hiện đại
        layout = QVBoxLayout(central_widget)
        layout.setContentsMargins(15, 15, 15, 15)
        
        # Khởi tạo Widget Tab để chứa các chức năng khác nhau
        self.tabs = QTabWidget()
        self.tabs.setFont(QFont("Segoe UI", 11, QFont.Bold))
        layout.addWidget(self.tabs)
        
        # Khởi tạo các Tab (mỗi tab là một module riêng biệt)
        self.tab1 = SplitExcelWidget()       # Tab 1: Xử lý tách danh sách nhân viên theo tổ
        self.tab2 = ScanProductionWidget()   # Tab 2: Scan ảnh bảng sản lượng bằng AI (Gemini)
        self.tab3 = LinkExcelWidget()        # Tab 3: Tích hợp dữ liệu JSON đã scan vào các file Excel
        
        # Thêm các tab vào TabWidget
        self.tabs.addTab(self.tab1, "1. Tách File Nhân Viên")
        self.tabs.addTab(self.tab2, "2. Scan Ảnh Sản Lượng")
        self.tabs.addTab(self.tab3, "3. Tích Hợp Excel")
        
        # Tùy chỉnh giao diện CSS cho các Tab (Tab Styling)
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
            QTabBar::tab:hover:!selected {
                background-color: #D1D5DB;
            }
        """)

if __name__ == "__main__":
    # Chạy thử giao diện nếu chạy file này trực tiếp
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec())

