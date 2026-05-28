import sys
from utils.logger import setup_logger

# Phải gọi setup_logger TRƯỚC khi import bất kỳ module nào khác có sử dụng logging
setup_logger()

from PySide6.QtWidgets import QApplication
from ui.gui import MainWindow
from utils.config import load_config
import debugpy

import ctypes
from PySide6.QtGui import QIcon

def main():
    debugpy.listen(("localhost", 5678))

    print("Debugger ready")
    
    # Thực hiện di chuyển cấu hình cũ sang thư mục ẩn
    try:
        from utils.paths import migrate_configs_to_hidden_dir
        migrate_configs_to_hidden_dir()
    except Exception as e:
        print(f"⚠️ Lỗi khi chạy migration cấu hình: {e}")
    
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
    app.setWindowIcon(QIcon("app_icon.png"))
    
    # Khởi tạo và hiển thị giao diện người dùng chính
    window = MainWindow()
    window.showMaximized()  # Hiển thị cửa sổ ở trạng thái maximize
    
    # Bắt đầu vòng lặp sự kiện của ứng dụng
    sys.exit(app.exec())

if __name__ == "__main__":
    main()

