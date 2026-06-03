"""
Tạo hệ thống chấm điểm tự động cho bài thi
- CLO 1: Câu 1-40 (mỗi câu 0.25 điểm, tổng 10 điểm)
- CLO 2: Câu 41-55 (mỗi câu 0.5 điểm, tổng 7.5 điểm)
"""

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

def create_scoring_system():
    # Tạo workbook mới
    wb = openpyxl.Workbook()
    
    # ========== SHEET 1: CONFIG ==========
    ws_config = wb.active
    ws_config.title = "Config"
    
    # Header
    ws_config['A1'] = "CẤU HÌNH HỆ THỐNG CHẤM ĐIỂM"
    ws_config['A1'].font = Font(bold=True, size=14)
    ws_config.merge_cells('A1:D1')
    
    # Phần 1: Định nghĩa CLO
    ws_config['A3'] = "ĐỊNH NGHĨA CHUẨN ĐẦU RA (CLO)"
    ws_config['A3'].font = Font(bold=True, size=11)
    
    headers = ['CLO', 'Tên CLO', 'Câu từ', 'Câu đến', 'Điểm/câu', 'Tổng điểm']
    for col, header in enumerate(headers, 1):
        cell = ws_config.cell(row=4, column=col)
        cell.value = header
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill(start_color="366092", end_color="366092", fill_type="solid")
    
    # Dữ liệu CLO
    clo_data = [
        ['CLO 1', 'Kiến thức cơ bản', 1, 40, 0.25, 10],
        ['CLO 2', 'Ứng dụng nâng cao', 41, 55, 0.5, 7.5]
    ]
    
    for row_idx, data in enumerate(clo_data, 5):
        for col_idx, value in enumerate(data, 1):
            ws_config.cell(row=row_idx, column=col_idx).value = value
    
    # Phần 2: Quản lý mã đề
    ws_config['A8'] = "QUẢN LÝ ĐÁP ÁN THEO MÃ ĐỀ"
    ws_config['A8'].font = Font(bold=True, size=11)
    
    headers2 = ['Mã đề', 'Câu 1', 'Câu 2', 'Câu 3', '...', 'Câu 55']
    for col, header in enumerate(headers2, 1):
        cell = ws_config.cell(row=9, column=col)
        cell.value = header
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill(start_color="70AD47", end_color="70AD47", fill_type="solid")
    
    # Ví dụ mã đề
    example_answers = ['A', 'B', 'C', 'D', 'A'] + ['B'] * 50  # Ví dụ 55 câu
    row_data = ['ĐHXD-2024-A'] + example_answers
    for col_idx, value in enumerate(row_data, 1):
        ws_config.cell(row=10, column=col_idx).value = value
    
    row_data = ['ĐHXD-2024-B'] + ['C', 'D', 'A', 'B', 'C'] + ['A'] * 50
    for col_idx, value in enumerate(row_data, 1):
        ws_config.cell(row=11, column=col_idx).value = value
    
    # Điều chỉnh độ rộng cột
    ws_config.column_dimensions['A'].width = 15
    ws_config.column_dimensions['B'].width = 12
    ws_config.column_dimensions['C'].width = 12
    
    # ========== SHEET 2: CHẤM ĐIỂM ==========
    ws_score = wb.create_sheet("Chấm Điểm")
    
    # Header
    ws_score['A1'] = "BẢNG CHẤM ĐIỂM BÀI THI"
    ws_score['A1'].font = Font(bold=True, size=14)
    ws_score.merge_cells('A1:F1')
    
    # Thông tin sinh viên
    ws_score['A3'] = "Mã sinh viên:"
    ws_score['B3'] = ""
    ws_score['A4'] = "Họ tên:"
    ws_score['B4'] = ""
    ws_score['A5'] = "Mã đề:"
    ws_score['B5'] = ""
    
    # Bảng chấm điểm
    headers3 = ['STT', 'Câu', 'Đáp án đúng', 'Đáp án SV', 'Đúng/Sai', 'Điểm']
    for col, header in enumerate(headers3, 1):
        cell = ws_score.cell(row=7, column=col)
        cell.value = header
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
    
    # Tạo dòng cho 55 câu
    for i in range(1, 56):
        row = 7 + i
        ws_score.cell(row=row, column=1).value = i  # STT
        ws_score.cell(row=row, column=2).value = f"Câu {i}"
        
        # Công thức kiểm tra đúng/sai
        ws_score.cell(row=row, column=5).value = f'=IF(C{row}=D{row},"✓","✗")'
        
        # Công thức tính điểm
        if i <= 40:
            ws_score.cell(row=row, column=6).value = f'=IF(E{row}="✓",0.25,0)'
        else:
            ws_score.cell(row=row, column=6).value = f'=IF(E{row}="✓",0.5,0)'
    
    # Tổng kết
    summary_row = 7 + 56
    ws_score.cell(row=summary_row, column=1).value = "TỔNG KẾT"
    ws_score.cell(row=summary_row, column=1).font = Font(bold=True)
    
    ws_score.cell(row=summary_row, column=5).value = "Tổng điểm:"
    ws_score.cell(row=summary_row, column=5).font = Font(bold=True)
    ws_score.cell(row=summary_row, column=6).value = f'=SUM(F8:F62)'
    ws_score.cell(row=summary_row, column=6).font = Font(bold=True, size=12)
    ws_score.cell(row=summary_row, column=6).fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
    
    # Điểm CLO
    clo1_row = summary_row + 2
    ws_score.cell(row=clo1_row, column=5).value = "Điểm CLO 1 (Câu 1-40):"
    ws_score.cell(row=clo1_row, column=5).font = Font(bold=True)
    ws_score.cell(row=clo1_row, column=6).value = f'=SUM(F8:F47)'
    ws_score.cell(row=clo1_row, column=6).font = Font(bold=True)
    ws_score.cell(row=clo1_row, column=6).fill = PatternFill(start_color="C6E0B4", end_color="C6E0B4", fill_type="solid")
    
    clo2_row = clo1_row + 1
    ws_score.cell(row=clo2_row, column=5).value = "Điểm CLO 2 (Câu 41-55):"
    ws_score.cell(row=clo2_row, column=5).font = Font(bold=True)
    ws_score.cell(row=clo2_row, column=6).value = f'=SUM(F48:F62)'
    ws_score.cell(row=clo2_row, column=6).font = Font(bold=True)
    ws_score.cell(row=clo2_row, column=6).fill = PatternFill(start_color="C6E0B4", end_color="C6E0B4", fill_type="solid")
    
    # Điều chỉnh độ rộng cột
    ws_score.column_dimensions['A'].width = 8
    ws_score.column_dimensions['B'].width = 10
    ws_score.column_dimensions['C'].width = 15
    ws_score.column_dimensions['D'].width = 15
    ws_score.column_dimensions['E'].width = 12
    ws_score.column_dimensions['F'].width = 12
    
    # ========== SHEET 3: HƯỚNG DẪN ==========
    ws_guide = wb.create_sheet("Hướng Dẫn")
    
    guide_content = [
        ["HƯỚNG DẪN SỬ DỤNG HỆ THỐNG CHẤM ĐIỂM"],
        [],
        ["BƯỚC 1: CHUẨN BỊ ĐÁP ÁN"],
        ["- Vào sheet 'Config'"],
        ["- Nhập đáp án đúng cho mã đề (dòng 10, 11, ...)"],
        ["- Mỗi mã đề 1 dòng, 55 cột cho 55 câu"],
        [],
        ["BƯỚC 2: NHẬP THÔNG TIN SINH VIÊN"],
        ["- Vào sheet 'Chấm Điểm'"],
        ["- Điền: Mã sinh viên, Họ tên, Mã đề"],
        [],
        ["BƯỚC 3: NHẬP ĐÁP ÁN SINH VIÊN"],
        ["- Ở cột 'Đáp án đúng' (C), copy từ sheet Config"],
        ["- Ở cột 'Đáp án SV' (D), dán đáp án của sinh viên"],
        ["- Các công thức sẽ tự động tính điểm"],
        [],
        ["LƯU Ý:"],
        ["- CLO 1 (Câu 1-40): 0.25 điểm/câu, tổng 10 điểm"],
        ["- CLO 2 (Câu 41-55): 0.5 điểm/câu, tổng 7.5 điểm"],
        ["- Tổng điểm tối đa: 17.5 điểm"],
        ["- Chỉnh sửa điểm/câu trong sheet 'Config' nếu cần"],
    ]
    
    for row_idx, row_data in enumerate(guide_content, 1):
        for col_idx, value in enumerate(row_data, 1):
            cell = ws_guide.cell(row=row_idx, column=col_idx)
            cell.value = value
            if row_idx == 1:
                cell.font = Font(bold=True, size=12)
    
    ws_guide.column_dimensions['A'].width = 50
    
    # Lưu file
    wb.save('scoring-system.xlsx')
    print("✅ File 'scoring-system.xlsx' đã tạo thành công!")
    print("📍 Vị trí: ./scoring-system.xlsx")
    print("\n📋 Cấu trúc file:")
    print("   - Sheet 'Config': Quản lý đáp án theo mã đề")
    print("   - Sheet 'Chấm Điểm': Nhập đáp án sinh viên và tính điểm")
    print("   - Sheet 'Hướng Dẫn': Hướng dẫn sử dụng")

if __name__ == "__main__":
    # Kiểm tra openpyxl
    try:
        import openpyxl
    except ImportError:
        print("⚠️ Cần cài đặt openpyxl:")
        print("pip install openpyxl")
        exit(1)
    
    create_scoring_system()
