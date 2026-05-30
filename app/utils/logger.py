import logging
import os
from app.utils.paths import get_log_path

# Thiết lập hệ thống ghi nhật ký (Logging) cho toàn bộ ứng dụng.
# Nhật ký sẽ được lưu vào file data/03_Logs/app.log và in ra màn hình console.
def setup_logger():
    log_dir = get_log_path()
    log_file = os.path.join(log_dir, 'app.log')
    
    # Đảm bảo thư mục log tồn tại
    if not os.path.exists(log_dir):
        os.makedirs(log_dir, exist_ok=True)
        
    # Xóa các handler cũ nếu có để tránh xung đột
    for handler in logging.root.handlers[:]:
        logging.root.removeHandler(handler)
        
    logging.basicConfig(
        level=logging.INFO,
        format='%(asctime)s - %(levelname)s - %(message)s',
        handlers=[
            logging.FileHandler(log_file, encoding='utf-8'),
            logging.StreamHandler()
        ],
        force=True # Ép buộc sử dụng cấu hình này, bỏ qua các cấu hình trước đó
    )
    logging.info(f"Hệ thống Log đã được khởi tạo tại: {log_file}")
