import asyncio
import os
import re
import pandas as pd
import PyPDF2
from playwright.async_api import async_playwright

def rename_pdf_from_content(filepath):
    try:
        with open(filepath, "rb") as f:
            reader = PyPDF2.PdfReader(f)
            text = ""
            for page in reader.pages:
                text += page.extract_text() + "\n"
        
        matches = re.findall(r'[A-Z]{4}\d{7}', text)
        if matches:
            unique_conts = list(dict.fromkeys(matches))
            cont_str = "_".join(unique_conts[:3])
            if len(unique_conts) > 3:
                cont_str += "_etc"
            
            new_filename = f"{cont_str}.pdf"
            new_filepath = os.path.join(os.path.dirname(filepath), new_filename)
            
            # Tránh trùng tên
            idx = 1
            while os.path.exists(new_filepath) and new_filepath != filepath:
                new_filepath = os.path.join(os.path.dirname(filepath), f"{cont_str}_{idx}.pdf")
                idx += 1
                
            os.rename(filepath, new_filepath)
            return os.path.basename(new_filepath)
    except Exception as e:
        print(f"    [!] Lỗi khi đọc nội dung PDF để đổi tên: {e}")
    return os.path.basename(filepath)

def get_search_values():
    url = 'https://docs.google.com/spreadsheets/d/1GFQYGt3b4D_tSzMWYcGwEiz-wbMaFJQKIQmG_p4QTBs/export?format=xlsx&gid=471286597'
    
    print("Đang tải dữ liệu từ Google Sheets...")
    df = pd.read_excel(url)
    
    fortune_vals = []
    
    # Tìm tháng ở cột A (NOTE)
    month_str = "Thang_5_2026"
    for val in df.iloc[:, 0]:
        if isinstance(val, pd.Timestamp):
            month_str = f"Thang_{val.month}_{val.year}"
            break
            
    if 'NƠI HẠ  CONT' in df.columns and 'SỐ CONT' in df.columns:
        mask = df['NƠI HẠ  CONT'].astype(str).str.contains(r'b[iì]nh\s*d[uư][oơ]ng|bdp', case=False, na=False)
        valid_rows = df[mask]
        for _, row in valid_rows.iterrows():
            cont = str(row['SỐ CONT']).strip()
            if cont and cont != 'nan':
                fortune_vals.append({'val': cont, 'month': month_str})
                
    return {"Fortune": fortune_vals, "HASITC": []}

async def process_company(page, company, search_items, auth_token):
    for item in search_items:
        val = item['val']
        month_str = item['month']
        
        print(f"[{company}] Đang tra cứu dữ liệu cho: {val} (Tháng: {month_str})")
        
        # Tạo thư mục theo công ty và tháng
        safe_month = "".join(c for c in month_str if c.isalnum() or c in " _-")
        if not safe_month.strip() or safe_month.lower() == "nan":
            safe_month = "Khac"
        target_dir = os.path.join("invoices", company, safe_month)
        os.makedirs(target_dir, exist_ok=True)
        
        # 2. Tải hóa đơn trực tiếp bằng Số Container
        print(f"  -> Đang tra cứu hóa đơn cho Số Cont {val}...")
        await page.goto("https://smartport.gemadept.com.vn/tracking/tracking_inv")
        await page.wait_for_timeout(1500)
        
        import re
        try:
            search_input = page.get_by_placeholder(re.compile(r"Nhập mã tra cứu", re.IGNORECASE))
            await search_input.fill(val)
        except Exception:
            # Fallback nếu placeholder thay đổi
            await page.locator("input.ant-input").first.fill(val)
            
        await page.locator("button.ant-input-search-button").click()
        
        try:
            await page.locator(".anticon-download").first.wait_for(timeout=10000)
        except:
            print(f"    [!] Không tìm thấy hóa đơn hoặc nút tải cho {val}")
            continue
            
        # Đếm số lượng hóa đơn
        download_buttons = await page.locator(".anticon-download").all()
        for idx, btn in enumerate(download_buttons):
            try:
                async with page.expect_download(timeout=15000) as download_info:
                    await btn.click()
                download = await download_info.value
                
                filename = f"temp_{val}_{idx+1}.pdf"
                filepath = os.path.join(target_dir, filename)
                await download.save_as(filepath)
                
                # Đổi tên file dựa vào nội dung (số cont)
                final_name = rename_pdf_from_content(filepath)
                
                print(f"    [OK] Đã tải và lưu thành: {final_name} (Thư mục: {target_dir})")
            except Exception as e:
                print(f"    [!] Lỗi tải {val} - dòng {idx+1}: {e}")

async def main():
    company_data = get_search_values()
    
    async with async_playwright() as p:
        browser = await p.chromium.launch(headless=False)
        context = await browser.new_context(accept_downloads=True)
        page = await context.new_page()
        
        await page.goto("https://smartport.gemadept.com.vn/")
        print("\n" + "="*50)
        print("Trình duyệt đã mở. Vui lòng đăng nhập và nhập Captcha!")
        print("Sau khi đăng nhập thành công, script sẽ tự động chạy.")
        print("="*50 + "\n")
        
        # Bắt token từ request
        auth_token = None
        
        async def handle_request(request):
            nonlocal auth_token
            try:
                auth = await request.header_value("authorization")
                if auth:
                    auth_token = auth
            except Exception:
                pass
                
        page.on("request", handle_request)
        
        # Chờ đăng nhập xong
        while True:
            if auth_token:
                print("Phát hiện đã đăng nhập (đã lấy được token)!")
                break
            await page.wait_for_timeout(1000)
            
        await page.wait_for_timeout(3000)
        
        # Bắt đầu xử lý
        for company, vals in company_data.items():
            print(f"\n--- Bắt đầu xử lý {company} ({len(vals)} items) ---")
            await process_company(page, company, vals, auth_token)
            
        print("\nHOÀN THÀNH TẤT CẢ!")
        await page.wait_for_timeout(5000)
        await browser.close()

if __name__ == "__main__":
    asyncio.run(main())
