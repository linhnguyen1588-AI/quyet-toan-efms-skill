import os
import pandas as pd

pdf_dir = r"D:\workspace-ai\CSHT\CSHT HAOHUA\PDFs"
excel_path = r"D:\workspace-ai\CSHT\CSHT HAOHUA\BANG KE CSHT HAOHUA T5_Done.xlsx"

pdf_files = os.listdir(pdf_dir)
pdf_tks = set()
for f in pdf_files:
    if f.endswith('.pdf'):
        tk = f.replace("THONG BÁO PHI CSHT _ ", "").replace(".pdf", "").strip()
        pdf_tks.add(tk)

df = pd.read_excel(excel_path, header=None)

# Find target row and cols
target_row = -1
tk_col = -1
tb_col = -1

for row in range(len(df)):
    for col in range(len(df.columns)):
        val = str(df.iloc[row, col]).strip().upper()
        if val == 'SỐ TỜ KHAI':
            target_row = row
            tk_col = col
        elif val == 'SỐ TB PHÍ':
            tb_col = col
    if target_row != -1 and tk_col != -1 and tb_col != -1:
        break

excel_tks_with_tb = set()
for row in range(target_row + 1, len(df)):
    tk = str(df.iloc[row, tk_col]).strip()
    if tk == 'nan' or tk == '': continue
    tb = str(df.iloc[row, tb_col]).strip()
    # Check if there is an actual fee notice number
    if tb != 'nan' and tb != '' and tb != '000000000000' and tb != 'None':
        excel_tks_with_tb.add(tk)

print(f"Tổng số PDF hiện có: {len(pdf_tks)}")
print(f"Tổng số tờ khai CÓ số TB phí trong Excel: {len(excel_tks_with_tb)}")

missing_in_pdf = excel_tks_with_tb - pdf_tks
missing_in_excel = pdf_tks - excel_tks_with_tb

if missing_in_pdf:
    print(f"CÓ trong Excel nhưng thiếu file PDF: {missing_in_pdf}")
else:
    print("Không có file PDF nào bị thiếu so với Excel.")

if missing_in_excel:
    print(f"CÓ file PDF nhưng thiếu trong Excel (chưa lưu số): {missing_in_excel}")
else:
    print("Không có số TB phí nào thiếu trong Excel so với PDF.")
