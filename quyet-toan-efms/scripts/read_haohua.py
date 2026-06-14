import pandas as pd
import sys

sys.stdout.reconfigure(encoding='utf-8')

file_path = r"D:\workspace-ai\CSHT\CSHT HAOHUA\BANG KE CSHT HAOHUA T5.xlsx"
try:
    df = pd.read_excel(file_path, sheet_name=0)
    df = df.dropna(how='all')
    print(df.head(20).to_string())
except Exception as e:
    print(f"Error reading file: {e}")
