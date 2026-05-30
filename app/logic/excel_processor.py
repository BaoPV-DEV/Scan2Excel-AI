import os
import logging

from app.logic.tab1_classifier import read_employee_data, classify_employees
from app.logic.tab1_writer import write_all_groups


# =========================================================
# TAB 1 ENGINE - PHÂN TÁCH DANH SÁCH NHÂN VIÊN THEO TỔ
# =========================================================
def process_excel(
    input_path,
    template_path,
    output_dir,
    mm,
    yyyy,
    log_callback,
    progress_callback=None
):
    """
    Pipeline chính:
    INPUT Excel → READ → CLASSIFY → VALIDATE → WRITE OUTPUT
    """

    mm_yyyy = f"{mm}/{yyyy}"

    # ============================
    # LOG START
    # ============================
    log_callback("=" * 55)
    log_callback("🚀 BẮT ĐẦU QUY TRÌNH TÁCH FILE NHÂN VIÊN")
    log_callback(f"   Tháng/Năm: {mm_yyyy}")
    log_callback(f"   File input: {os.path.basename(input_path)}")
    log_callback(f"   Output: {output_dir}")
    log_callback("=" * 55)

    try:
        # ============================
        # STEP 1: READ DATA
        # ============================
        if progress_callback:
            progress_callback(5)

        df = read_employee_data(input_path, log_callback)

        if df is None or len(df) == 0:
            raise ValueError("File input không có dữ liệu hợp lệ")

        # ============================
        # STEP 2: CLASSIFY EMPLOYEES
        # ============================
        if progress_callback:
            progress_callback(15)

        classified = classify_employees(df, mm, log_callback)

        # ============================
        # STEP 2.5: VALIDATION CHECK
        # ============================
        total_input = len(df)

        total_classified = sum(
            len(group.get("employees", []))
            for group in classified.values()
        )

        if total_classified != total_input:
            missing = total_input - total_classified
            log_callback(
                f"\n⚠️ CẢNH BÁO: Thiếu {missing} nhân viên trong phân loại!"
            )
        else:
            log_callback(
                f"\n✅ OK: {total_classified}/{total_input} nhân viên đã phân loại đầy đủ"
            )

        # ============================
        # STEP 3: CREATE OUTPUT DIR
        # ============================
        os.makedirs(output_dir, exist_ok=True)

        # ============================
        # STEP 4: WRITE OUTPUT FILES
        # ============================
        if progress_callback:
            progress_callback(20)

        success, message = write_all_groups(
            classified,
            output_dir,
            mm,
            yyyy,
            log_callback,
            progress_callback
        )

        # ============================
        # STEP 5: SUMMARY
        # ============================
        log_callback("=" * 55)
        log_callback(f"✅ {message}")
        log_callback(f"📂 Output: {output_dir}")
        log_callback("=" * 55)

        logging.info(f"Tab1 completed: {message}")

        if progress_callback:
            progress_callback(100)

        return success, message

    # =========================================================
    # ERROR HANDLING (CHUẨN HOÁ MESSAGE)
    # =========================================================
    except FileNotFoundError as e:
        log_callback(f"❌ File không tồn tại: {e}")
        return False, f"Không tìm thấy file:\n{str(e)}"

    except PermissionError as e:
        log_callback(f"❌ File đang bị khóa: {e}")
        return False, f"File đang mở hoặc bị khóa:\n{str(e)}"

    except ValueError as e:
        log_callback(f"❌ Dữ liệu không hợp lệ: {e}")
        return False, f"Lỗi dữ liệu:\n{str(e)}"

    except Exception as e:
        log_callback(f"❌ Lỗi hệ thống: {e}")
        logging.error("Tab1 error", exc_info=True)
        return False, f"Lỗi không xác định:\n{str(e)}"