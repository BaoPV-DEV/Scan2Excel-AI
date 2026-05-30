import os

# Tải các thiết lập cấu hình cơ bản cho ứng dụng.
# Hiện tại chủ yếu chứa các đường dẫn mặc định và định dạng file.
def load_config():
    config = {
        'model_paths': {
            'paddleocr': 'models/'  # Thư mục chứa các model OCR (nếu dùng PaddleOCR)
        },
        'output_directory': 'output/', # Thư mục xuất file mặc định (legacy)
        'default_settings': {
            'image_format': 'png'       # Định dạng ảnh ưu tiên khi xử lý
        }
    }
    return config
