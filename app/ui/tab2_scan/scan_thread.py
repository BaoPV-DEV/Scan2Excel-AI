import os
import json
import re
import PIL.Image
from datetime import datetime
from google import genai
from PySide6.QtCore import QThread, Signal
from app.utils.paths import get_json_data_path

# ==============================================================================
# Luồng chạy ngầm gửi ảnh đơn lẻ lên Google Gemini AI (Giữ lại để tương thích ngược)
# ==============================================================================
class ProcessingThread(QThread):
    progress = Signal(int)
    log = Signal(str)
    finished = Signal(dict)

    def __init__(self, image_paths, api_key, model_name="gemini-3.1-flash-lite"):
        super().__init__()
        self.image_paths = image_paths
        self.api_key = api_key
        self.model_name = model_name

    def run(self):
        try:
            if not self.api_key:
                raise ValueError("API Key không hợp lệ. Vui lòng kiểm tra file .env")

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
            
            response = client.models.generate_content(
                model=self.model_name,
                contents=contents
            )
            
            if not response or not response.text:
                raise Exception("AI không trả về nội dung.")

            json_match = re.search(r'\{.*\}', response.text, re.DOTALL)
            if json_match:
                final_json = json.loads(json_match.group())
                self.finished.emit(final_json)
            else:
                raise Exception("Không trích xuất được dữ liệu JSON.")

        except Exception as e:
            self.log.emit(f"❌ Lỗi: {str(e)}")
            self.finished.emit({})


# ==============================================================================
# Luồng chạy ngầm Quét Batch OCR Hàng Loạt Tuần Tự (Sử dụng cho hệ thống mới)
# ==============================================================================
class BatchProcessingThread(QThread):
    item_started = Signal(int, str)  # row_idx, subdir_name
    item_finished = Signal(int, str, dict, str)  # row_idx, "success"/"failed", data_dict, saved_json_path/error_msg
    progress = Signal(int)  # Tiến độ tổng quan (%)
    log = Signal(str)  # Log gửi về giao diện
    finished = Signal()  # Báo hoàn tất toàn bộ tiến trình

    def __init__(self, batch_data, api_key, model_name="gemini-3.1-flash-lite"):
        """
        batch_data: list of dicts:
        [
            {
                "row": int,
                "subdir_name": str,
                "subdir_path": str,
                "images": list of image absolute paths,
                "bonus": int
            }
        ]
        """
        super().__init__()
        self.batch_data = batch_data
        self.api_key = api_key
        self.model_name = model_name

    def run(self):
        try:
            if not self.api_key:
                raise ValueError("API Key không hợp lệ. Vui lòng cấu hình API Key.")

            client = genai.Client(api_key=self.api_key)
            total = len(self.batch_data)
            
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

            for idx, item in enumerate(self.batch_data):
                row = item["row"]
                subdir_name = item["subdir_name"]
                images = item["images"]
                bonus = item["bonus"]
                
                self.item_started.emit(row, subdir_name)
                self.log.emit(f"⏳ [{idx+1}/{total}] Đang quét mã hàng: {subdir_name} ({len(images)} ảnh)...")
                
                try:
                    images_to_send = []
                    for p in images:
                        if os.path.exists(p):
                            images_to_send.append(PIL.Image.open(p))
                    
                    if not images_to_send:
                        raise Exception("Không tìm thấy tệp ảnh nào hợp lệ.")
                    
                    contents = [prompt] + images_to_send
                    
                    response = client.models.generate_content(
                        model=self.model_name,
                        contents=contents
                    )
                    
                    if not response or not response.text:
                        raise Exception("AI không phản hồi dữ liệu.")
                    
                    json_match = re.search(r'\{.*\}', response.text, re.DOTALL)
                    if not json_match:
                        raise Exception("Không trích xuất được dữ liệu JSON từ AI.")
                    
                    data = json.loads(json_match.group())
                    
                    # Áp dụng mức thưởng mã hàng mới cấu hình từ UI
                    if "thong_tin_chung" in data:
                        data["thong_tin_chung"]["thuong_ma_hang_moi"] = bonus
                        if not data["thong_tin_chung"].get("ma_hang"):
                            data["thong_tin_chung"]["ma_hang"] = subdir_name
                    
                    # Lưu JSON tự động
                    json_path = self.save_json(data, subdir_name)
                    if not json_path:
                        raise Exception("Ghi file JSON thất bại.")
                    
                    self.item_finished.emit(row, "success", data, json_path)
                    
                except Exception as e:
                    self.log.emit(f"❌ Quét thất bại Mã hàng {subdir_name}: {e}")
                    self.item_finished.emit(row, "failed", {}, str(e))
                
                # Cập nhật tiến độ (%)
                progress_pct = int(((idx + 1) / total) * 100)
                self.progress.emit(progress_pct)
                
                # Nghỉ trễ giữa các mã tránh nghẽn API (2.5 giây)
                if idx < total - 1:
                    import time
                    time.sleep(2.5)

            self.finished.emit()

        except Exception as e:
            self.log.emit(f"❌ Lỗi hệ thống Batch: {str(e)}")
            self.finished.emit()

    def save_json(self, data, subdir_name):
        try:
            thong_tin = data.get("thong_tin_chung", {})
            thoi_gian = thong_tin.get("thoi_gian", "N/A")
            to_sx = thong_tin.get("to_san_xuat", "N/A")
            ma_hang = thong_tin.get("ma_hang", subdir_name)
            
            match_date = re.search(r'(\d{1,2})[/-](\d{4})', thoi_gian)
            month = match_date.group(1).zfill(2) if match_date else datetime.now().strftime("%m")
            year = match_date.group(2) if match_date else datetime.now().strftime("%Y")
            
            to_val = str(to_sx).strip()
            team_num_match = re.search(r'\d+', to_val)
            team_folder = f"Tổ {team_num_match.group()}" if team_num_match else "Tổ Unknown"
            
            json_dir = get_json_data_path(year, month, team_folder)
            c_ma_hang = re.sub(r'[\\/*?:"<>|]', "_", str(ma_hang)).strip() or "N_A"
            json_path = os.path.join(json_dir, f"{c_ma_hang}.json")
            
            with open(json_path, 'w', encoding='utf-8') as f:
                json.dump(data, f, indent=4, ensure_ascii=False)
            return json_path
        except Exception:
            return None
