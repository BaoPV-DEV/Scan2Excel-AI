import os
import gc
import pythoncom
import win32com.client as win32
from dotenv import load_dotenv

# Hàm chuyển đổi cột số thành chữ cái (VD: 6 -> "F", 29 -> "AC")
def get_column_letter(col_idx):
    """Chuyển đổi chỉ số cột (1-based) thành chữ cái Excel."""
    result = ""
    while col_idx > 0:
        col_idx, remainder = divmod(col_idx - 1, 26)
        result = chr(65 + remainder) + result
    return result

def update_templates_with_headers(titles, log_callback, progress_callback=None):
    """
    Cập nhật danh sách tiêu đề cột F -> AC ở dòng 4 sheet 'Bang TH nop' cho 2 template:
    - Template\\sx\\to_may_template.xlsx
    - Template\\sx\\kiem_hoa_template.xlsx
    """
    load_dotenv()
    EXCEL_PASS = os.getenv("EXCEL_SHEET_PASSWORD", "8863")

    # Xác định đường dẫn tuyệt đối của các template
    logic_dir = os.path.abspath(os.path.dirname(__file__))
    workspace_dir = os.path.abspath(os.path.join(logic_dir, "../.."))
    
    template_paths = {
        "Tổ May Template": os.path.join(workspace_dir, "resources", "templates", "sx", "to_may_template.xlsx"),
        "Kiểm Hóa Template": os.path.join(workspace_dir, "resources", "templates", "sx", "kiem_hoa_template.xlsx")
    }

    # Kiểm tra sự tồn tại của file
    for name, path in template_paths.items():
        if not os.path.exists(path):
            log_callback(f"❌ Không tìm thấy file template '{name}' tại: {path}")
            return False, f"Thiếu file: {os.path.basename(path)}"

    # Giới hạn số lượng tiêu đề cột từ F (6) đến AC (29) -> tối đa 24 cột
    MAX_COLS = 24
    if len(titles) > MAX_COLS:
        log_callback(f"⚠️ Danh sách tiêu đề ({len(titles)}) vượt quá giới hạn 24 cột (F -> AC). Chỉ lấy 24 cột đầu tiên.")
        titles = titles[:MAX_COLS]

    pythoncom.CoInitialize()
    excel_app = None
    success_count = 0

    try:
        log_callback("🚀 Khởi động ứng dụng Excel...")
        excel_app = win32.DispatchEx("Excel.Application")
        excel_app.Visible = False
        excel_app.DisplayAlerts = False
        excel_app.ScreenUpdating = False

        total_templates = len(template_paths)
        for idx, (name, path) in enumerate(template_paths.items()):
            log_callback(f"\n📂 Đang mở template: {os.path.basename(path)}")
            wb = None
            try:
                wb = excel_app.Workbooks.Open(os.path.abspath(path), UpdateLinks=0, ReadOnly=False, Password=EXCEL_PASS)
            except Exception as e:
                log_callback(f"⚠️ Không thể mở bằng mật khẩu, thử mở bình thường: {e}")
                try:
                    wb = excel_app.Workbooks.Open(os.path.abspath(path))
                except Exception as ex:
                    log_callback(f"❌ Lỗi mở file {os.path.basename(path)}: {ex}")
                    continue

            try:
                # Tìm sheet "Bang TH nop"
                ws = None
                for sheet in wb.Sheets:
                    if sheet.Name == "Bang TH nop":
                        ws = sheet
                        break

                if not ws:
                    log_callback(f"❌ Không tìm thấy Sheet 'Bang TH nop' trong template: {os.path.basename(path)}")
                    continue

                # Mở khóa sheet nếu bị khóa
                try:
                    ws.Unprotect(Password=EXCEL_PASS)
                except:
                    try:
                        ws.Unprotect()
                    except:
                        pass

                # Xóa trắng nội dung cũ trong khoảng F4:AC4 để giữ nguyên format
                log_callback("🧼 Đang xóa sạch nội dung tiêu đề cũ từ F4 đến AC4...")
                ws.Range("F4:AC4").ClearContents()

                # Điền tiêu đề và ẩn/hiện cột tương ứng
                log_callback("✍️ Đang ghi tiêu đề mới và thiết lập ẩn/hiện cột...")
                for col in range(6, 30):  # F (6) -> AC (29)
                    title_idx = col - 6
                    col_letter = get_column_letter(col)

                    if title_idx < len(titles):
                        val = titles[title_idx].strip()
                        ws.Cells(4, col).Value = val
                        ws.Columns(col).EntireColumn.Hidden = False
                        log_callback(f"   🔹 Cột {col_letter}: Hiện thị & Ghi giá trị '{val}'")
                    else:
                        ws.Columns(col).EntireColumn.Hidden = True
                        log_callback(f"   💤 Cột {col_letter}: Ẩn cột (Không dùng)")



                # Lưu và đóng
                wb.Save()
                log_callback(f"💾 Đã lưu thành công: {os.path.basename(path)}")
                success_count += 1

            except Exception as e:
                log_callback(f"❌ Lỗi khi cập nhật sheet của file {os.path.basename(path)}: {e}")
            finally:
                if wb:
                    try:
                        wb.Close(SaveChanges=True)
                    except:
                        pass
                    del wb

            # Cập nhật tiến độ
            if progress_callback:
                progress_callback(int(((idx + 1) / total_templates) * 100))

    except Exception as e:
        log_callback(f"❌ Lỗi hệ thống: {e}")
        return False, str(e)
    finally:
        if excel_app:
            try:
                excel_app.ScreenUpdating = True
                excel_app.DisplayAlerts = True
                excel_app.Quit()
            except:
                pass
            del excel_app
        pythoncom.CoUninitialize()
        gc.collect()

    if success_count == len(template_paths):
        log_callback("\n🎉 CẬP NHẬT CẢ 2 TEMPLATE HOÀN TẤT THÀNH CÔNG!")
        return True, "Cập nhật thành công 2 template."
    else:
        log_callback(f"\n⚠️ Hoàn tất với cảnh báo: Cập nhật thành công {success_count}/{len(template_paths)} template.")
        return False, f"Chỉ cập nhật thành công {success_count}/{len(template_paths)} template."
