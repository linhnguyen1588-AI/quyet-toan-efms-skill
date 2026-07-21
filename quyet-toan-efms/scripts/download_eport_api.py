import asyncio
import os
import sys
import pandas as pd
import openpyxl
from openpyxl.styles import PatternFill
from datetime import datetime, timedelta
from playwright.async_api import async_playwright

sys.stdout.reconfigure(encoding='utf-8', line_buffering=True)

def resolve_excel_path(user_input):
    user_input = user_input.strip()
    if not user_input:
        user_input = "HOA DON EPORT CAT LAI 01.07.xlsx"

    search_dirs = [
        r"D:\workspace-ai\HOA DON EPORT CAT LAI",
        r"D:\workspace-ai",
        r"C:\Users\sonha\workspace-ai\HOA DON EPORT CAT LAI",
        r"C:\Users\sonha\workspace-ai"
    ]

    candidates = [
        user_input,
        user_input + ".xlsx",
        user_input + ".xls"
    ]

    for d in search_dirs:
        for c in candidates:
            p = os.path.join(d, c)
            if os.path.exists(p):
                return os.path.abspath(p)

    for c in candidates:
        if os.path.exists(c):
            return os.path.abspath(c)

    search_base = r"D:\workspace-ai"
    if os.path.exists(search_base):
        for root, _, files in os.walk(search_base):
            for f in files:
                if f.lower().endswith(('.xlsx', '.xls')):
                    if user_input.lower() in f.lower() or user_input.lower() in os.path.splitext(f)[0].lower():
                        return os.path.abspath(os.path.join(root, f))

    return user_input

async def wait_if_captcha(page):
    while await page.locator("text=Vui lòng nhập mã này").count() > 0 or await page.locator("text=What code is in the image").count() > 0:
        print("\n[!] PHÁT HIỆN CAPTCHA! Đang đứng chờ giải captcha...")
        await page.wait_for_timeout(4000)

async def main():
    raw_input = sys.argv[1].strip() if len(sys.argv) > 1 and sys.argv[1].strip() else "HOA DON EPORT CAT LAI 01.07.xlsx"
    excel_path = resolve_excel_path(raw_input)
        
    out_dir = os.path.join(os.path.dirname(excel_path), "HOA DON")
    if not os.path.exists(out_dir):
        os.makedirs(out_dir, exist_ok=True)
        
    print(f"[SMART SCAN] Đã tìm thấy đường dẫn file Excel: {excel_path}")
    print(f"[SMART SCAN] Thư mục lưu file Hóa Đơn PDF: {out_dir}")

    try:
        df = pd.read_excel(excel_path)
        booking_col = None
        for col in df.columns:
            col_str = str(col).strip().upper()
            if 'BOOK' in col_str or 'BILL' in col_str or 'CONT' in col_str:
                booking_col = col
                break
        if booking_col:
            print(f"-> Đã phát hiện cột chứa mã: '{booking_col}'")
            bookings = df[booking_col].dropna().astype(str).unique().tolist()
        else:
            print("-> Dùng cột đầu tiên trong file Excel...")
            bookings = df.iloc[:, 0].dropna().astype(str).unique().tolist()
            
        bookings = [b.strip() for b in bookings if str(b).strip() and str(b).strip().lower() != 'nan']
        target_books = bookings
    except Exception as e:
        print("Lỗi đọc file Excel:", e)
        return

    if not target_books:
        print("Không tìm thấy số Book/Bill nào trong file!")
        return

    print(f"Sẽ tiến hành quét tổng cộng {len(target_books)} mã...")
    failed_books = []
    success_count = 0

    async with async_playwright() as p:
        # Headless background mode
        browser = await p.chromium.launch(headless=True)
        context = await browser.new_context(accept_downloads=True)
        page = await context.new_page()
        
        print("\n==================================================")
        print("SIÊU ROBOT TẢI HÓA ĐƠN ePORT CÁT LÁI CHẠY NGẦM 100%")
        print("Tự động đăng nhập ngầm & quét tải hóa đơn PDF...")
        print("==================================================\n")
        
        await page.goto("https://eport.saigonnewport.com.vn/Home/Login")
        
        # Tự động đăng nhập
        try:
            await page.wait_for_timeout(2000)
            await page.fill("input[name='Username'], input[id='Username']", "0314436809")
            await page.fill("input[name='Password'], input[id='Password']", "Sotrans1234@")
            
            login_btn = page.locator("input[type='submit'], button[type='submit'], .btn-login, button:visible").first
            await login_btn.click()
            print("[INFO] Đã gửi thông tin đăng nhập tài khoản 0314436809 vào ePort Tân Cảng...")
            await page.wait_for_timeout(4000)
            await wait_if_captcha(page)
        except Exception as e:
            print(f"[WARNING] Đăng nhập tự động: {e}")
        
        one_month_ago = (datetime.now() - timedelta(days=60)).strftime("%d/%m/%Y")
        
        for idx, book in enumerate(target_books, 1):
            print(f"\n[{idx}/{len(target_books)}] >>> Đang xử lý Book/Bill: {book}")
            try:
                await page.goto("https://eport.saigonnewport.com.vn/FullContainerDelivery")
                await wait_if_captcha(page)
                await page.wait_for_selector("input, textarea", timeout=10000)
                
                inputs = await page.locator("input, textarea").all()
                for inp in inputs:
                    ph = await inp.get_attribute("placeholder") or ""
                    name = await inp.get_attribute("name") or ""
                    if "từ ngày" in ph.lower() or "fromdate" in name.lower():
                        await inp.fill("")
                        await inp.fill(one_month_ago)
                        
                found_note = False
                for inp in inputs:
                    ph = await inp.get_attribute("placeholder") or ""
                    name = await inp.get_attribute("name") or ""
                    if "ghi ch" in ph.lower() or "note" in name.lower() or "remark" in name.lower():
                        await inp.fill(book)
                        found_note = True
                        break
                        
                if not found_note:
                    await page.evaluate(f'''() => {{
                        let inputs = document.querySelectorAll('input, textarea');
                        for (let el of inputs) {{
                            let html = el.outerHTML.toLowerCase();
                            if (html.includes('ghichu') || html.includes('note') || html.includes('remark')) {{
                                el.value = '{book}';
                                el.dispatchEvent(new Event('input', {{ bubbles: true }}));
                                break;
                            }}
                        }}
                    }}''')
                    
                btns = await page.locator("button, a").all()
                for btn in btns:
                    text = await btn.text_content() or ""
                    if "tìm kiếm" in text.lower() or "lấy dữ liệu" in text.lower():
                        await btn.click()
                        break
                        
                await page.wait_for_timeout(3000)
                
                grid = page.locator(".dx-datagrid, table").first
                inv_locators = [
                    "a[title*='Hóa đơn']", 
                    "a[title*='hóa đơn']", 
                    "i.fa-file-invoice", 
                    "i.fa-download"
                ]
                
                downloaded = False
                for loc in inv_locators:
                    elements = await grid.locator(loc).all()
                    if elements:
                        print(f"  -> Đã tìm thấy {len(elements)} hóa đơn. Đang tải về thư mục...")
                        for el in elements:
                            try:
                                async with page.expect_download(timeout=10000) as download_info:
                                    await el.click(force=True)
                                    download = await download_info.value
                                    
                                    file_path = os.path.join(out_dir, f"{book}.pdf")
                                    counter = 1
                                    while os.path.exists(file_path):
                                        file_path = os.path.join(out_dir, f"{book}_{counter}.pdf")
                                        counter += 1
                                        
                                    await download.save_as(file_path)
                                    print(f"  [SUCCESS] ĐÃ TẢI LƯU THÀNH CÔNG FILE PDF: {file_path}")
                                    downloaded = True
                                    success_count += 1
                            except Exception as e:
                                pass
                        
                        if downloaded:
                            break 
                        
                if not downloaded:
                    print(f"  [!] Book {book} không có dữ liệu hóa đơn hoặc đã quá hạn trên ePort.")
                    failed_books.append(book)
            except Exception as e:
                print(f"  [!] Lỗi khi xử lý Book {book}: {e}")
                failed_books.append(book)
                
        print(f"\n🎉 HOÀN THÀNH ROBOT TẢI HÓA ĐƠN NGẦM! TẢI THÀNH CÔNG {success_count} FILE PDF.")
        await browser.close()
        
    if failed_books:
        print(f"\n[+] Đang tô vàng {len(failed_books)} Book không có hóa đơn vào file Excel...")
        try:
            wb = openpyxl.load_workbook(excel_path)
            ws = wb.active
            booking_col_idx = None
            for row in ws.iter_rows(min_row=1, max_row=5):
                for cell in row:
                    if cell.value and ("BOOKING" in str(cell.value).upper() or "BILL" in str(cell.value).upper()):
                        booking_col_idx = cell.column
                        break
                if booking_col_idx: break
            
            if booking_col_idx:
                yellow_fill = PatternFill(start_color="FFFFFF00", end_color="FFFFFF00", fill_type="solid")
                for row in ws.iter_rows(min_row=2):
                    cell_val = str(row[booking_col_idx - 1].value).strip() if row[booking_col_idx - 1].value else ""
                    if cell_val in failed_books:
                        for cell in row:
                            cell.fill = yellow_fill
                
                try:
                    wb.save(excel_path)
                    print("  [v] Đã tô vàng xong file Excel!")
                except PermissionError:
                    print("  [!] Vui lòng đóng file Excel để hệ thống lưu tô màu.")
        except Exception as e:
            pass

if __name__ == "__main__":
    asyncio.run(main())
