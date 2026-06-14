import asyncio
import re
import os
import pandas as pd
from playwright.async_api import async_playwright

SHEET_ID = "1GFQYGt3b4D_tSzMWYcGwEiz-wbMaFJQKIQmG_p4QTBs"
GID = "471286597"
CSV_URL = f"https://docs.google.com/spreadsheets/d/{SHEET_ID}/export?format=csv&gid={GID}"

def get_search_values():
    url = 'https://docs.google.com/spreadsheets/d/1GFQYGt3b4D_tSzMWYcGwEiz-wbMaFJQKIQmG_p4QTBs/export?format=xlsx&gid=471286597'
    try:
        print("Đang tải dữ liệu từ Google Sheets...")
        df = pd.read_excel(url)
        
        fortune_vals = []
        if 'NƠI HẠ  CONT' in df.columns and 'SỐ CONT' in df.columns:
            mask = df['NƠI HẠ  CONT'].astype(str).str.contains(r'bình\s*dương|binh\s*duong|bdp', case=False, na=False)
            valid_rows = df[mask]
            for _, row in valid_rows.iterrows():
                cont = str(row['SỐ CONT']).strip()
                if cont and cont != 'nan':
                    fortune_vals.append(cont)
                    
        return {"Fortune": fortune_vals}
    except Exception as e:
        print(f"Lỗi khi đọc Google Sheets: {e}")
        return {}

async def main():
    company_data = get_search_values()
    if not company_data:
        return
        
    pincode_map = {}
    
    async with async_playwright() as p:
        browser = await p.chromium.launch(headless=False)
        context = await browser.new_context(accept_downloads=True)
        page = await context.new_page()
        
        print("\n==================================================")
        print("Trình duyệt đã mở. Vui lòng đăng nhập!")
        print("==================================================\n")
        
        await page.goto("https://smartport.gemadept.com.vn/")
        
        auth_token = None
        async def handle_request(request):
            nonlocal auth_token
            try:
                auth = await request.header_value("authorization")
                if auth:
                    auth_token = auth
            except:
                pass
        page.on("request", handle_request)
        
        while not auth_token:
            await page.wait_for_timeout(1000)
            
        print("\n>>> GIAI ĐOẠN 1: TÌM PINCODE <<<")
        print("1. Bấm menu 'Tra cứu Container' (hoặc Tra cứu e-Eir).")
        print("2. Chọn Cảng Bình Dương.")
        print("3. QUAN TRỌNG: Chọn khoảng thời gian (Từ ngày - Đến ngày) bao trùm tất cả các tháng cần lấy (VD: 01/01 đến 31/12) để tránh lỗi không tìm thấy.")
        print("4. Tự tra cứu 1 container bất kỳ làm mẫu.")
        print("-> Khi bảng kết quả hiện ra, robot sẽ tự chạy tiếp 126 cont còn lại!")
        
        await page.wait_for_selector("table, .ant-table", timeout=0)
        await page.wait_for_timeout(2000)
        
        headers = await page.locator("th").all_inner_texts()
        pincode_idx = -1
        for i, header in enumerate(headers):
            if "pincode" in header.lower() or "pin code" in header.lower():
                pincode_idx = i
                break
                
        if pincode_idx == -1:
            pincode_idx = 4
            
        print(f"Bắt đầu tra cứu PinCode (Cột {pincode_idx+1})")
        
        for company, vals in company_data.items():
            for val in vals:
                try:
                    search_input = page.get_by_placeholder(re.compile(r"nhập.*container", re.IGNORECASE)).first
                    if await search_input.count() == 0:
                        search_input = page.locator("input[type='text']").first
                    
                    await search_input.fill("")
                    await search_input.fill(val)
                    await page.keyboard.press("Enter")
                    await page.wait_for_timeout(1000)
                    
                    try:
                        row = page.locator("tbody tr").first
                        pincode = await row.locator("td").nth(pincode_idx).inner_text(timeout=1500)
                        pincode = pincode.strip()
                        
                        if pincode and pincode != "-":
                            pincode_map[val] = pincode
                            print(f"[OK] {val} -> {pincode}")
                        else:
                            print(f"[!] Không có Pin cho {val}")
                    except Exception:
                        print(f"[!] Không có dữ liệu cho {val}")
                        
                except Exception as e:
                    print(f"[!] Lỗi tra Pin cho {val}: {e}")
                    
        print("\n>>> GIAI ĐOẠN 2: TẢI HÓA ĐƠN <<<")
        print("-> Sếp bấm sang menu 'Tra cứu Hóa đơn điện tử' đi ạ!")
        
        while "tracking_inv" not in page.url:
            await page.wait_for_timeout(1000)
            
        print("Đã vào trang Tra cứu hóa đơn. Bắt đầu tải...")
        await page.wait_for_timeout(2000)
        
        for cont, pin in pincode_map.items():
            try:
                search_input = page.get_by_placeholder(re.compile(r"Số Order|PinCode|Nhập mã tra cứu", re.IGNORECASE)).first
                if await search_input.count() == 0:
                    search_input = page.locator("input[type='text']").first
                    
                await search_input.fill("")
                await search_input.fill(pin)
                await page.keyboard.press("Enter")
                
                btn = page.locator(".anticon-download").first
                await btn.wait_for(timeout=10000)
                
                async with page.expect_download() as download_info:
                    await btn.click()
                download = await download_info.value
                
                folder_path = os.path.join(os.path.dirname(__file__), "..", "invoices", "Fortune", "Thang_5_2026")
                os.makedirs(folder_path, exist_ok=True)
                pdf_path = os.path.join(folder_path, f"{cont}.pdf")
                await download.save_as(pdf_path)
                print(f"  [v] Đã lưu: {cont}.pdf")
                
            except Exception as e:
                print(f"  [!] Lỗi hóa đơn {cont}: {e}")

        print("\n🎉 HOÀN THÀNH!")
        await page.wait_for_timeout(5000)
        await browser.close()

if __name__ == "__main__":
    asyncio.run(main())
