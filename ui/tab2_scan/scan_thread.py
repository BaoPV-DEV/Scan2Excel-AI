import os
import json
import re
import PIL.Image
from google import genai
from PySide6.QtCore import QThread, Signal

# Luồng chạy ngầm gửi ảnh lên Google Gemini AI và nhận về dữ liệu dạng JSON.
class ProcessingThread(QThread):
    progress = Signal(int)
    log = Signal(str)
    finished = Signal(dict)

    # Khởi tạo luồng và nhận dữ liệu đầu vào
    def __init__(self, image_paths, api_key, model_name="gemini-3.1-flash-lite"):
        super().__init__()
        self.image_paths = image_paths
        self.api_key = api_key
        self.model_name = model_name

    # Thực hiện tiến trình quét ảnh bằng AI
    def run(self):
        try:
            if not self.api_key:
                raise ValueError("API Key không hợp lệ. Vui lòng kiểm tra file .env")

            # 1. Khởi tạo Client Gemini AI
            client = genai.Client(api_key=self.api_key)
            
            prompt = """
            Hãy đóng vai trò là một chuyên gia kiểm soát sản xuất ngành may. 
            Phân tích các hình ảnh bảng sản lượng đính kèm và trả về JSON DUY NHẤT.
            Cấu trúc JSON yêu cầu:
            {
              "thong_tin_chung": {
                "don_vi": "S&D Co., Ltd",
                "loai_san_pham": "string",
                "ma_hang": "string",
                "khach_hang": "string",
                "to_san_xuat": "string",
                "thoi_gian": "string",
                "thuong_ma_hang_moi": int,
                "tong_san_luong_muc_tieu": int
              },
              "danh_sach_cong_doan": [
                {
                  "stt": int,
                  "mo_ta": "string",
                  "dinh_muc_t": float,
                  "thuc_hien": [
                    {"nhan_vien": "string hoặc N/A", "ma_nhan_vien": "string 5 số", "so_luong": int}
                  ],
                  "tong_thuc_hien": int
                }
              ]
            }
            """

            self.log.emit(f"🚀 Đang gửi ảnh đến AI (Model: {self.model_name}) để phân tích...")
            
            images_to_send = []
            for path in self.image_paths:
                if os.path.exists(path):
                    images_to_send.append(PIL.Image.open(path))

            if not images_to_send:
                raise Exception("Không tìm thấy tệp ảnh nào hợp lệ.")

            contents = [prompt] + images_to_send
            
            # Gọi API Gemini để xử lý nội dung
            response = client.models.generate_content(
                model=self.model_name,
                contents=contents
            )
            
            if not response or not response.text:
                raise Exception("AI không trả về nội dung.")

            # Trích xuất JSON từ kết quả văn bản
            json_match = re.search(r'\{.*\}', response.text, re.DOTALL)
            if json_match:
                final_json = json.loads(json_match.group())
                self.finished.emit(final_json)
            else:
                raise Exception("Không trích xuất được dữ liệu JSON.")

        except Exception as e:
            self.log.emit(f"❌ Lỗi: {str(e)}")
            self.finished.emit({})
