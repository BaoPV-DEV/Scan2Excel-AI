import os
import logging
from app.logic.tab1_classifier import read_employee_data, classify_employees
from app.logic.tab1_writer import write_all_groups

# ==============================================================================
# Hàm chính xử lý tách file Excel danh sách nhân viên theo tổ/bộ phận
# Quy trình: Đọc file input -> Phân loại NV -> Ghi vào template -> Lưu file
# ==============================================================================
def process_excel(input_path, template_path, output_dir, mm, yyyy, log_callback, progress_callback=None):
    """
    Hàm xử lý chính cho Tab 1: Phân chia danh sách nhân viên.
    
    Tham số:
        input_path: Đường dẫn file Excel danh sách NV gốc
        template_path: (Legacy - không sử dụng, template được chọn tự động)
        output_dir: Thư mục lưu kết quả (ROOT_PATH/01_danh_sach_chia_to/YYYY/MM)
        mm: Tháng (định dạng 2 chữ số, VD: "04")
        yyyy: Năm (định dạng 4 chữ số, VD: "2026")
        log_callback: Hàm ghi nhật ký ra giao diện
        progress_callback: Hàm cập nhật tiến độ (0-100)
    """
    mm_yyyy = f"{mm}/{yyyy}"
    
    log_callback(f"{'='*55}")
    log_callback(f"🚀 BẮT ĐẦU QUY TRÌNH TÁCH FILE NHÂN VIÊN")
    log_callback(f"   Tháng/Năm: {mm_yyyy}")
    log_callback(f"   File đầu vào: {os.path.basename(input_path)}")
    log_callback(f"   Thư mục lưu: {output_dir}")
    log_callback(f"{'='*55}")

    try:
        # ============================================================
        # BƯỚC 1: Đọc và làm sạch dữ liệu từ file input
        # ============================================================
        if progress_callback:
            progress_callback(5)
        df = read_employee_data(input_path, log_callback)

        # ============================================================
        # BƯỚC 2: Phân loại nhân viên theo quy tắc ưu tiên
        # (Thời vụ -> Tổ 1-12/Tổ trưởng -> Bộ phận đặc thù -> Chưa xử lý)
        # ============================================================
        if progress_callback:
            progress_callback(15)
        classified = classify_employees(df, mm, log_callback)

        # ============================================================
        # BƯỚC 2.5: Kiểm chứng - đảm bảo không bỏ sót nhân viên nào
        # ============================================================
        total_input = len(df)
        total_classified = sum(len(info["employees"]) for info in classified.values())
        if total_classified != total_input:
            log_callback(f"\n⚠️ CẢNH BÁO: Tổng NV đầu vào ({total_input}) ≠ Tổng đã phân loại ({total_classified})!")
            log_callback(f"   Có {total_input - total_classified} nhân viên bị bỏ sót!")
        else:
            log_callback(f"\n✅ Kiểm chứng OK: {total_classified}/{total_input} nhân viên đã được phân loại đầy đủ.")

        # ============================================================
        # BƯỚC 3: Tạo thư mục output nếu chưa tồn tại
        # ============================================================
        os.makedirs(output_dir, exist_ok=True)

        # ============================================================
        # BƯỚC 4: Ghi dữ liệu vào các file Excel theo template tương ứng
        # ============================================================
        if progress_callback:
            progress_callback(20)
        success, message = write_all_groups(classified, output_dir, mm, yyyy, log_callback, progress_callback)

        # ============================================================
        # BƯỚC 5: Tổng kết
        # ============================================================
        log_callback(f"\n{'='*55}")
        log_callback(f"✅ {message}")
        log_callback(f"📂 Thư mục lưu trữ: {output_dir}")
        log_callback(f"{'='*55}")

        logging.info(f"Tab1 - Tách file hoàn tất: {message}")
        return success, message

    except FileNotFoundError as e:
        log_callback(f"❌ Lỗi: Không tìm thấy file - {str(e)}")
        return False, f"Không tìm thấy file:\n{str(e)}"
    except ValueError as e:
        log_callback(f"❌ Lỗi dữ liệu: {str(e)}")
        return False, f"Lỗi dữ liệu:\n{str(e)}"
    except PermissionError as e:
        log_callback(f"❌ Lỗi quyền truy cập: {str(e)}")
        return False, f"Lỗi quyền truy cập (File đang mở):\n{str(e)}"
    except Exception as e:
        log_callback(f"❌ Lỗi không xác định: {str(e)}")
        logging.error(f"Tab1 - Lỗi: {str(e)}", exc_info=True)
        return False, f"Đã xảy ra lỗi:\n{str(e)}"
