import win32com.client
import pythoncom
import config, glob, os

pythoncom.CoInitialize()

path = None
for ext in ("*.xlsx", "*.xls", "*.xlsm"):
    matches = glob.glob(os.path.join(config.IKSAN_FILE_DIR, "익산대장*" + ext.lstrip("*")))
    if matches:
        path = matches[0]
        break

print(f"파일: {path}")

xl = win32com.client.Dispatch("Excel.Application")
xl.Visible = False
xl.DisplayAlerts = False
wb = xl.Workbooks.Open(os.path.abspath(path))
ws = wb.Worksheets(1)

# 열 자체의 기본 색상 확인
col_color = int(ws.Columns(12).Interior.Color)
cr = col_color & 0xFF; cg = (col_color >> 8) & 0xFF; cb = (col_color >> 16) & 0xFF
print(f"\nL열(12) 컬럼 기본 Interior.Color = #{cr:02X}{cg:02X}{cb:02X}")
print(f"L열 컬럼 Pattern = {ws.Columns(12).Interior.Pattern}")

# 특정 행 확인 (1~5행 + 620~625행 + 마지막 몇 행)
print("\n=== 샘플 셀 색상 ===")
sample_rows = list(range(1, 6)) + list(range(618, 626)) + [1750, 1752, 1754]
for row in sample_rows:
    cell = ws.Cells(row, 12)
    val = cell.Value
    ic = int(cell.Interior.Color)
    r = ic & 0xFF; g = (ic >> 8) & 0xFF; b = (ic >> 16) & 0xFF
    print(f"  행{row:4d} | #{r:02X}{g:02X}{b:02X} | 값={str(val)[:40] if val else '(빈셀)'}")

wb.Close(False)
xl.Quit()
pythoncom.CoUninitialize()


wb.Close(False)
xl.Quit()
pythoncom.CoUninitialize()
