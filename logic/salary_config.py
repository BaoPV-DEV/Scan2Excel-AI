# ==============================================================================
# SALARY DEPLOYMENT CONFIGURATION - Tab 5
# Linking configurations for salary template deployment with external data sources
# ==============================================================================
# Các placeholder hỗ trợ:
#   {row}: Dòng hiện tại (VD: 9)
#   {row_minus_1}: Dòng hiện tại trừ 1 (VD: 8)
#   {row_plus_1}: Dòng hiện tại cộng 1 (VD: 10)
#   {yyyy}: Năm hiện tại (VD: 2026)
#   {yy_short}: Hai chữ số cuối của năm hiện tại (VD: 26)
#   {mm}: Tháng hiện tại dạng 2 chữ số (VD: 04)
#   {mm_int}: Tháng hiện tại dạng số nguyên (VD: 4)
#   {prev_year}: Năm trước (VD: 2025)
#   {prev_yy_short}: Hai chữ số cuối của năm trước (VD: 25)
#   {prev_month}: Tháng trước dạng 2 chữ số (VD: 03)
#   {prev_month_int}: Tháng trước dạng số nguyên (VD: 3)
# ==============================================================================

# External linking formulas for salary templates
SALARY_TEMPLATE_FORMULA_CONFIG = {
    "gt/bao_ve_template.xlsx": {
        "Bang TH nop": {
            "start_row": 9,
            "row_step": 2,
            "formulas": {
                "D": "=+IFERROR(VLOOKUP(C{row},'D:\\HỒ SƠ+LƯƠNG SD\\LƯƠNG\\LƯƠNG + CÔNG {yyyy}\\Tháng {mm}-{yyyy}\\Danh sách+Chấm công\\[Danh sách CBCNV bản dùng làm lương T{mm_int}.xlsx]Danh sách'!$D$1:$AS$1000,42,0),0)"
            }
        },
        "Lương {mm}": {
            "start_row": 10,
            "row_step": 2,
            "formulas": {
                "Z": "=+IFERROR(VLOOKUP(C{row_minus_1},'D:\\HỒ SƠ+LƯƠNG SD\\LƯƠNG\\LƯƠNG + CÔNG {prev_year}\\Tháng {prev_month}-{prev_year}\\[CK tháng {prev_month}.{prev_year}.xlsx]Bản gốc T{prev_month}'!$C$1:$J$1000,8,0),0)",
                "AA": "=+IFERROR(VLOOKUP(C{row_minus_1},'D:\\HỒ SƠ+LƯƠNG SD\\LƯƠNG\\LƯƠNG + CÔNG {prev_year}\\Tháng {prev_month}-{prev_year}\\[CK tháng {prev_month}.{prev_year}.xlsx]Bản gốc T{prev_month}'!$C$1:$F$1000,4,0),0)"
            }
        }
    },
    "gt/bep_an_va_cong_vu_template.xlsx": {
        "Bang TH nop": {
            "start_row": 9,
            "row_step": 2,
            "formulas": {
                "D": "=+IFERROR(VLOOKUP(C{row},'D:\\HỒ SƠ+LƯƠNG SD\\LƯƠNG\\LƯƠNG + CÔNG {yyyy}\\Tháng {mm}-{yyyy}\\Danh sách+Chấm công\\[Danh sách CBCNV bản dùng làm lương T{mm_int}.xlsx]Danh sách'!$D$1:$AS$1000,42,0),0)",
                "E": "=+IFERROR(VLOOKUP(C{row},'D:\\HỒ SƠ+LƯƠNG SD\\LƯƠNG\\LƯƠNG + CÔNG {yyyy}\\Tháng {mm}-{yyyy}\\Danh sách+Chấm công\\[Danh sách CBCNV làm lương tháng {mm}.xlsx]Danh Sách dùng'!$E$5:$AG$11000,29,0),0)",
                "G": "=+IFERROR(VLOOKUP(C{row},'D:\\HỒ SƠ+LƯƠNG SD\\LƯƠNG\\LƯƠNG + CÔNG {yyyy}\\Tháng {mm}-{yyyy}\\Danh sách+Chấm công\\[Chấm công {mm}.{yy_short}.xlsx]Công'!$C$5:$AO$6000,39,0),0)",
                "H": "=+IFERROR(VLOOKUP(C{row},'D:\\HỒ SƠ+LƯƠNG SD\\LƯƠNG\\LƯƠNG + CÔNG {yyyy}\\Tháng {mm}-{yyyy}\\Danh sách+Chấm công\\[Chấm công {mm}.{yy_short}.xlsx]Công'!$C$5:$AU$6000,44,0),0)",
                "I": "=+IFERROR(VLOOKUP(C{row},'D:\\HỒ SƠ+LƯƠNG SD\\LƯƠNG\\LƯƠNG + CÔNG {yyyy}\\Tháng {mm}-{yyyy}\\Danh sách+Chấm công\\[Chấm công {mm}.{yy_short}.xlsx]Công'!$C$5:$AR$6000,42,0),0)",
                "K": "=+IFERROR(VLOOKUP(C{row},'D:\\HỒ SƠ+LƯƠNG SD\\LƯƠNG\\LƯƠNG + CÔNG {yyyy}\\Tháng {mm}-{yyyy}\\Danh sách+Chấm công\\[Chấm công {mm}.{yy_short}.xlsx]Công'!$C$5:$BC$6000,53,0),0)",
                "L": "=+IFERROR(VLOOKUP(C{row},'D:\\HỒ SƠ+LƯƠNG SD\\LƯƠNG\\LƯƠNG + CÔNG {yyyy}\\Tháng {mm}-{yyyy}\\Danh sách+Chấm công\\[Chấm công {mm}.{yy_short}.xlsx]Công'!$C$5:$BE$6000,55,0),0)",
                "M": "=+IFERROR(VLOOKUP(C{row},'D:\\HỒ SƠ+LƯƠNG SD\\LƯƠNG\\LƯƠNG + CÔNG {yyyy}\\Tháng {mm}-{yyyy}\\Danh sách+Chấm công\\[phụ cấp con nhỏ năm {yyyy}.xlsx]Con nhỏ'!$A$3:$B$3000,2,0),0)",
                "R": "=+IFERROR(VLOOKUP(C{row},'D:\\HỒ SƠ+LƯƠNG SD\\LƯƠNG\\LƯƠNG + CÔNG {yyyy}\\Tháng {mm}-{yyyy}\\Danh sách+Chấm công\\[phụ cấp con nhỏ năm {yyyy}.xlsx]Nhóm 6'!$C$2:$E$2000,3,0),0)",
                "T": "=+IFERROR(VLOOKUP(C{row},'D:\\HỒ SƠ+LƯƠNG SD\\LƯƠNG\\LƯƠNG + CÔNG {yyyy}\\Tháng {mm}-{yyyy}\\Danh sách+Chấm công\\[phụ cấp con nhỏ năm {yyyy}.xlsx]Thâm niên'!$B$3:$G$3000,6,0),0)",
                "U": "=+IFERROR(VLOOKUP(C{row},'D:\\HỒ SƠ+LƯƠNG SD\\LƯƠNG\\LƯƠNG + CÔNG {yyyy}\\Tháng {mm}-{yyyy}\\Danh sách+Chấm công\\[Chấm công {mm}.{yy_short}.xlsx]Công'!$C$5:$AP$6000,40,0),0)"
            }
        },
        "Lương {mm}": {
            "start_row": 10,
            "row_step": 2,
            "formulas": {
                "H": "=+(F{row}/26/8*{bep_cong_vu_rate})*H{row_minus_1}",
                "Y": "=+IFERROR(VLOOKUP(C{row_minus_1},'D:\\HỒ SƠ+LƯƠNG SD\\LƯƠNG\\LƯƠNG + CÔNG {prev_year}\\Tháng {prev_month}-{prev_year}\\[CK tháng {prev_month}.{prev_year}.xlsx]Bản gốc T{prev_month}'!$C$1:$J$1000,8,0),0)",
                "Z": "=+IFERROR(VLOOKUP(C{row_minus_1},'D:\\HỒ SƠ+LƯƠNG SD\\LƯƠNG\\LƯƠNG + CÔNG {prev_year}\\Tháng {prev_month}-{prev_year}\\[CK tháng {prev_month}.{prev_year}.xlsx]Bản gốc T{prev_month}'!$C$1:$F$1000,4,0),0)"
            }
        }
    }
}
