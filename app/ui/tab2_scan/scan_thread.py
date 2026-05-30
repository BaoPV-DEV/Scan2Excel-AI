import os
import json
import re
import time
from datetime import datetime
from contextlib import ExitStack

import PIL.Image
from google import genai

from PySide6.QtCore import QThread, Signal
from app.utils.paths import get_json_data_path


# =========================
# GLOBAL OPTIMIZATION
# =========================
_JSON_REGEX = re.compile(r"\{.*\}", re.DOTALL)

BASE_PROMPT = """
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
""".strip()


def extract_json(text: str):
    """Trích JSON an toàn từ response AI"""
    if not text:
        return None
    match = _JSON_REGEX.search(text)
    if not match:
        return None
    try:
        return json.loads(match.group())
    except json.JSONDecodeError:
        return None


def load_images(image_paths):
    """Load ảnh an toàn + tự close tài nguyên"""
    images = []
    for p in image_paths:
        if os.path.exists(p):
            try:
                img = PIL.Image.open(p)
                images.append(img.copy())
                img.close()
            except Exception:
                continue
    return images


# ======================================================================
# SINGLE IMAGE THREAD (Backward compatible)
# ======================================================================
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

            self.log.emit(f"🚀 Đang gửi ảnh đến AI (Model: {self.model_name}) để phân tích...")

            images = load_images(self.image_paths)
            if not images:
                raise Exception("Không tìm thấy tệp ảnh nào hợp lệ.")

            response = client.models.generate_content(
                model=self.model_name,
                contents=[BASE_PROMPT] + images
            )

            data = extract_json(response.text if response else "")
            if not data:
                raise Exception("Không trích xuất được JSON từ AI.")

            self.finished.emit(data)

        except Exception as e:
            self.log.emit(f"❌ Lỗi: {str(e)}")
            self.finished.emit({})


# ======================================================================
# BATCH PROCESSING THREAD (OPTIMIZED)
# ======================================================================
class BatchProcessingThread(QThread):
    item_started = Signal(int, str)
    item_finished = Signal(int, str, dict, str)
    progress = Signal(int)
    log = Signal(str)
    finished = Signal()

    def __init__(self, batch_data, api_key, model_name="gemini-3.1-flash-lite"):
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

            for idx, item in enumerate(self.batch_data):
                row = item["row"]
                subdir_name = item["subdir_name"]
                images = item["images"]
                bonus = item["bonus"]

                self.item_started.emit(row, subdir_name)
                self.log.emit(f"⏳ [{idx+1}/{total}] Đang quét: {subdir_name}...")

                try:
                    imgs = load_images(images)
                    if not imgs:
                        raise Exception("Không có ảnh hợp lệ.")

                    response = client.models.generate_content(
                        model=self.model_name,
                        contents=[BASE_PROMPT] + imgs
                    )

                    data = extract_json(response.text if response else "")
                    if not data:
                        raise Exception("AI không trả về JSON hợp lệ.")

                    # inject business rule (GIỮ NGUYÊN LOGIC)
                    thong_tin = data.get("thong_tin_chung", {})
                    thong_tin["thuong_ma_hang_moi"] = bonus
                    if not thong_tin.get("ma_hang"):
                        thong_tin["ma_hang"] = subdir_name

                    json_path = self.save_json(data, subdir_name)
                    if not json_path:
                        raise Exception("Không ghi được file JSON.")

                    self.item_finished.emit(row, "success", data, json_path)

                except Exception as e:
                    self.log.emit(f"❌ Lỗi {subdir_name}: {e}")
                    self.item_finished.emit(row, "failed", {}, str(e))

                # progress update
                self.progress.emit(int(((idx + 1) / total) * 100))

                if idx < total - 1:
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

            match_date = re.search(r"(\d{1,2})[/-](\d{4})", thoi_gian)
            month = match_date.group(1).zfill(2) if match_date else datetime.now().strftime("%m")
            year = match_date.group(2) if match_date else datetime.now().strftime("%Y")

            team_match = re.search(r"\d+", str(to_sx))
            team_folder = f"Tổ {team_match.group()}" if team_match else "Tổ Unknown"

            json_dir = get_json_data_path(year, month, team_folder)

            safe_name = re.sub(r'[\\/*?:"<>|]', "_", str(ma_hang)).strip() or "N_A"
            json_path = os.path.join(json_dir, f"{safe_name}.json")

            os.makedirs(json_dir, exist_ok=True)

            with open(json_path, "w", encoding="utf-8") as f:
                json.dump(data, f, ensure_ascii=False, indent=4)

            return json_path

        except Exception:
            return None