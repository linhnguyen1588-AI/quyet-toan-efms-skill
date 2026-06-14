import pandas as pd
import sys
sys.stdout.reconfigure(encoding='utf-8')
xls = pd.ExcelFile(r'D:\SOTRANS LINH 2\LINH\THANH TOAN 2026\THANG 5\Copy of GIAY TOAN LUC THANG 05-2026.xls', engine='calamine')
for s in xls.sheet_names:
    df = pd.read_excel(xls, sheet_name=s, header=None)
    text = df.astype(str).to_string().upper()
    if 'THÁNG 5' in text or '05/2026' in text or '2026-05' in text or 'THANG 5' in text or 'T05' in text or 'T5' in text:
        print(f"FOUND IN SHEET: {s}")
