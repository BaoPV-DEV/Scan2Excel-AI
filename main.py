import sys
from app.utils.logger import setup_logger

# Phải gọi setup_logger TRƯỚC khi import bất kỳ module nào khác có sử dụng logging
setup_logger()

from PySide6.QtWidgets import QApplication
from app.ui.gui import MainWindow
from app.utils.config import load_config, resource_path
import ctypes
from PySide6.QtGui import QIcon

def main():
    
    # Sửa lỗi Icon ở Taskbar trên Windows (để app có icon riêng thay vì icon Python mặc định)
    myappid = 'sd.scan2excel.ai.1' 
    try:
        ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(myappid)
    except Exception:
        pass
        
    # Tải cấu hình từ file .env hoặc config (nếu có)
    config = load_config()
    
    # Khởi tạo ứng dụng PySide6
    app = QApplication(sys.argv)
    
    # Thiết lập icon cho cửa sổ chính
    icon = QIcon(resource_path("resources/icons/app_icon.png"))
    app.setWindowIcon(icon)

    window = MainWindow()
    window.setWindowIcon(icon)
    window.showMaximized()  # Hiển thị cửa sổ ở trạng thái maximize
    
    # Bắt đầu vòng lặp sự kiện của ứng dụng
    sys.exit(app.exec())

if __name__ == "__main__":
    main()

