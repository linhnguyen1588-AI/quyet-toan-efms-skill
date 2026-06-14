import win32com.client as win32
import os

excel_path = r"D:\workspace-ai\HOA DON\HAOHUA\HAOHUA XUAT EPORT.xlsx"

try:
    excel = win32.Dispatch("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False

    wb = excel.Workbooks.Open(excel_path)
    ws = wb.Sheets(1)

    lastRow = ws.UsedRange.Rows.Count

    # Fill down Booking (Column B)
    for i in range(2, lastRow + 1):
        val = ws.Cells(i, 2).Value
        if val is None or str(val).strip() == "":
            ws.Cells(i, 2).Value = ws.Cells(i - 1, 2).Value

    # Sort Range by Booking
    rng = ws.Range(f"A1:I{lastRow}")
    # Sort key is B2 (column B)
    rng.Sort(Key1=ws.Range("B2"), Order1=1, Header=1)

    wb.Save()
    print("Success")
except Exception as e:
    print(f"Error: {e}")
finally:
    try:
        wb.Close()
        excel.Quit()
    except:
        pass
