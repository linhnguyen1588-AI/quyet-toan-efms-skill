import asyncio
import os
import pandas as pd
import openpyxl
from playwright.async_api import async_playwright

async def run_checker():
    # 1. Read the local Excel file that we downloaded
    file_path = "ToanLuc.xlsx"
    if not os.path.exists(file_path):
        print("Không tìm thấy file ToanLuc.xlsx")
        return
        
    df = pd.read_excel(file_path, header=None)
    
    # 2. Find the row for Month 6 of 2026
    # The user put a date like 2026-06-01 in Column B (Index 1) as a header/divider.
    thang6_row = -1
    for r in range(len(df)):
        val = str(df.iloc[r, 1]).strip()
        if '2026-06' in val:
            thang6_row = r
            break
            
    if thang6_row == -1:
        print("Không tìm thấy dấu hiệu tháng 6/2026 ở cột B")
        return
        
    print(f"-> Bắt đầu lấy dữ liệu từ dòng Excel {thang6_row + 1}")
    
    # We will process from thang6_row + 1 downwards. Stop if we see next month divider.
    declarations_to_check = []
    for r in range(thang6_row + 1, len(df)):
        val_b = str(df.iloc[r, 1]).strip()
        if '2026-07' in val_b:
            break # Reached next month
            
        tk = val_b
        date_str = str(df.iloc[r, 3]).strip()
        
        # Only process if TK is a number
        if tk and tk != 'nan' and len(tk) >= 10:
            # Format date to dd/mm/yyyy
            try:
                if ' ' in date_str: # 2023-05-26 00:00:00
                    date_obj = pd.to_datetime(date_str)
                    formatted_date = date_obj.strftime("%d/%m/%Y")
                else:
                    formatted_date = date_str
                declarations_to_check.append((r, tk, formatted_date))
            except Exception as e:
                print(f"Lỗi parse ngày cho TK {tk}: {e}")
                
    if not declarations_to_check:
        print("Không có tờ khai nào trong Tháng 6!")
        return
        
    print(f"-> Tổng cộng {len(declarations_to_check)} tờ khai cần kiểm tra.")

    # 3. Setup Playwright
    results = {}
    async with async_playwright() as p:
        browser = await p.chromium.launch(headless=False)
        context = await browser.new_context()
        page = await context.new_page()
        
        await page.goto("https://pus.customs.gov.vn/faces/ContainerBarcode", wait_until="networkidle")
        
        for r, tk, formatted_date in declarations_to_check:
            print(f"Đang tra cứu tờ khai: {tk} - {formatted_date}")
            
            # Remove any existing alerts first by clicking dong
            try:
                alert = await page.query_selector('.x1g8.af_messages_dialog')
                if alert and await alert.is_visible():
                    dong_btn = await alert.query_selector('a.xfp:has-text("Đồng ý"), a.xfp:has-text("Đóng")')
                    if dong_btn:
                        await dong_btn.click()
                        await asyncio.sleep(1)
            except:
                pass
                
            try:
                await page.fill('input[name="pt1:it1"]', "0305623305") # Doanh nghiệp
                await page.fill('input[name="pt1:it2"]', tk)          # Số tờ khai
                await page.fill('input[name="pt1:it3"]', "02CI")      # Hải quan
                
                # Input Date
                await page.fill('input[name="pt1:it4"]', "") # Clear first
                await page.fill('input[name="pt1:it4"]', formatted_date)
                
                # Click Lấy thông tin
                await page.click('a.xfp:has-text("Lấy thông tin")')
                
                # Wait for response
                await asyncio.sleep(2)
                
                # Check for Alert Popup
                alert = await page.query_selector('.x1g8.af_messages_dialog')
                status = "Chưa TQ"
                if alert and await alert.is_visible():
                    text = await alert.inner_text()
                    text_lower = text.lower()
                    if "đã được hq giám sát xác nhận" in text_lower or "đủ điều kiện" in text_lower:
                        status = "TQ"
                    elif "không tồn tại" in text_lower or "chưa" in text_lower or "lỗi" in text_lower:
                        status = "Chưa TQ"
                    else:
                        print(f"  -> Lỗi không rõ: {text.strip()}")
                        status = "Chưa TQ"
                        
                    # Close alert
                    try:
                        dong_btn = await alert.query_selector('a.xfp:has-text("Đồng ý"), a.xfp:has-text("Đóng")')
                        if dong_btn:
                            await dong_btn.click()
                            await asyncio.sleep(0.5)
                    except:
                        pass
                else:
                    # No alert, maybe success and table loaded
                    empty_txt = await page.query_selector('#pt1\\:t1\\:\\:emptyTxt')
                    if empty_txt and await empty_txt.is_visible():
                        status = "Chưa TQ"
                    else:
                        status = "TQ"
                        
                print(f"  -> Kết quả: {status}")
                results[r] = status
                
            except Exception as e:
                print(f"  -> Lỗi khi tra cứu: {e}")
                results[r] = "Lỗi"
                
        await browser.close()
        
    # 4. Save results back to Excel
    try:
        wb = openpyxl.load_workbook(file_path)
        ws = wb.active
        # In Pandas, row index r -> Excel row r + 1 (because pandas is 0-indexed)
        # Column H is the 8th column
        for r, status in results.items():
            ws.cell(row=r+1, column=8, value=status)
            
        out_path = file_path.replace(".xlsx", "_ThongQuan.xlsx")
        wb.save(out_path)
        print(f"Đã lưu kết quả ra file: {out_path}")
    except Exception as e:
        print(f"Lỗi khi lưu Excel: {e}")

if __name__ == "__main__":
    asyncio.run(run_checker())
