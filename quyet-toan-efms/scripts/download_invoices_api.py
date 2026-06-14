import asyncio
import re
import os
import json
import pandas as pd
from playwright.async_api import async_playwright

import io
import requests
import openpyxl

def get_all_visible_containers():
    url = 'https://docs.google.com/spreadsheets/d/1GFQYGt3b4D_tSzMWYcGwEiz-wbMaFJQKIQmG_p4QTBs/export?format=xlsx'
    try:
        print("Đang tải dữ liệu từ Google Sheets...")
        resp = requests.get(url)
        
        # Dùng openpyxl để lấy danh sách sheet không bị ẩn
        wb = openpyxl.load_workbook(io.BytesIO(resp.content), read_only=True)
        visible_sheets = [s.title for s in wb.worksheets if s.sheet_state == 'visible']
        
        # Dùng pandas để parse dữ liệu
        xl = pd.ExcelFile(io.BytesIO(resp.content))
        
        all_containers = []
        for sheet_name in visible_sheets:
            print(f"  -> Đang quét sheet: {sheet_name}")
            try:
                df = xl.parse(sheet_name)
                # Tìm các cột cần thiết (có thể có khoảng trắng dư thừa)
                # Chuẩn hóa tên cột
                df.columns = [str(c).strip().upper() for c in df.columns]
                
                # Tìm cột NƠI HẠ CONT và SỐ CONT
                col_noi_ha = next((c for c in df.columns if 'NƠI HẠ' in c), None)
                col_so_cont = next((c for c in df.columns if 'SỐ CONT' in c or 'CONTAINER' in c), None)
                col_loading = next((c for c in df.columns if 'LOADING' in c), None)
                
                if col_noi_ha and col_so_cont:
                    # Lọc cảng Bình Dương
                    mask = df[col_noi_ha].astype(str).str.contains(r'bình\s*dương|binh\s*duong|bdp', case=False, na=False)
                    valid_rows = df[mask]
                    
                    for _, row in valid_rows.iterrows():
                        cont = str(row[col_so_cont]).strip()
                        if cont and cont != 'nan':
                            from_date = "01/01/2025"
                            to_date = "31/12/2026"
                            if col_loading:
                                ld = pd.to_datetime(row[col_loading], errors='coerce')
                                if pd.notnull(ld):
                                    from_date = (ld - pd.Timedelta(days=30)).strftime("%d/%m/%Y")
                                    to_date = (ld + pd.Timedelta(days=30)).strftime("%d/%m/%Y")
                            
                            all_containers.append({
                                "val": cont,
                                "from_date": from_date,
                                "to_date": to_date,
                                "sheet_name": sheet_name
                            })
            except Exception as e:
                print(f"    [!] Lỗi khi quét sheet {sheet_name}: {e}")
                
        return all_containers
    except Exception as e:
        print(f"Lỗi đọc dữ liệu Google Sheets: {e}")
        return []

async def main():
    containers = get_all_visible_containers()
    if not containers:
        print("Không tìm thấy container nào trong các sheet hiển thị!")
        return
        
    print(f"Đã lấy {len(containers)} containers từ toàn bộ file!")
    
    # pincode_map giờ sẽ lưu theo dạng {cont: {"pin": pin, "sheet": sheet_name}}
    pincode_map = {}
    
    async with async_playwright() as p:
        browser = await p.chromium.launch(headless=False)
        context = await browser.new_context(accept_downloads=True)
        page = await context.new_page()
        
        print("\n==================================================")
        print("Trình duyệt đã mở. Vui lòng đăng nhập!")
        print("==================================================\n")
        
        await page.goto("https://smartport.gemadept.com.vn/")
        
        api_headers = {}
        api_payload = {}
        api_url = ""
        
        async def handle_request(request):
            nonlocal api_headers, api_payload, api_url
            if "lazyloadES" in request.url and request.method == "POST":
                try:
                    data = request.post_data
                    if data:
                        api_payload = json.loads(data)
                        api_headers = request.headers
                        api_url = request.url
                except Exception as e:
                    pass
                    
        page.on("request", handle_request)
        
        print("\n>>> GIAI ĐOẠN 1: LẤY API MẪU <<<")
        print("1. Bấm menu 'Tra cứu Container' (hoặc Tra cứu e-Eir).")
        print("2. Chọn Cảng Bình Dương.")
        print("3. Tự gõ 1 container rồi bấm Tìm kiếm.")
        print("-> Chỉ cần làm 1 lần, robot sẽ copy lại cấu trúc để bắn API ngầm cho toàn bộ phần còn lại cực siêu tốc!")
        
        # Đợi người dùng search thành công 1 lần
        while not api_payload:
            await page.wait_for_timeout(1000)
            
        print("\n[+] ĐÃ BẮT ĐƯỢC API MẪU! Bắt đầu cày ngầm...")
        
        # Bắn API cho tất cả containers
        api_context = context.request
        
        for item in containers:
            cont = item['val']
            sheet_name = item['sheet_name']
            payload = api_payload.copy()
            payload["kw"] = cont
            payload["FindValue"] = cont
            payload["fromDate"] = item["from_date"]  # DD/MM/YYYY
            payload["toDate"] = item["to_date"]
            
            # format cho FromDate/ToDate: YYYY-MM-DD HH:MM:SS
            d_f, m_f, y_f = item["from_date"].split("/")
            payload["FromDate"] = f"{y_f}-{m_f}-{d_f} 00:00:00"
            d_t, m_t, y_t = item["to_date"].split("/")
            payload["ToDate"] = f"{y_t}-{m_t}-{d_t} 23:59:59"
            
            try:
                # Gửi API ngầm
                response = await api_context.post(api_url, headers=api_headers, data=payload)
                if response.status == 200:
                    resp_json = await response.json()
                            
                    data_list = resp_json.get("payload", [])
                    if data_list:
                        first_item = data_list[0]
                        pin = first_item.get("PinCode")
                        if pin and pin != "-":
                            pincode_map[cont] = {"pin": pin, "sheet": sheet_name}
                            print(f"[OK] [{sheet_name}] {cont} -> {pin}")
                        else:
                            print(f"[!] [{sheet_name}] {cont} -> API trả về mảng nhưng không có PinCode")
                    else:
                        print(f"[!] [{sheet_name}] {cont} -> Không có dữ liệu")
                else:
                    print(f"[!] [{sheet_name}] {cont} -> Lỗi API {response.status}")
            except Exception as e:
                print(f"[!] {cont} -> Lỗi kết nối API: {e}")
                
        print("\n>>> GIAI ĐOẠN 2: TẢI HÓA ĐƠN (QUA API) <<<")
        terminal_code = api_payload.get("siteId", "BDP")
        
        for cont, data in pincode_map.items():
            pin = data["pin"]
            sheet_name = data["sheet"]
            
            # Tách tên công ty từ sheet_name (VD: "FORTUNE - BẢNG KÊ - T5-2026" -> "FORTUNE")
            company_name = sheet_name.split("-")[0].strip()
            
            # Tạo đường dẫn lưu trữ chuẩn: d:\workspace-ai\HOA DON\TÊN CÔNG TY\THÁNG 5
            # Lùi về đúng gốc thư mục workspace-ai nếu script nằm trong workspace
            workspace_dir = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", "..", ".."))
            month_dir = os.path.join(workspace_dir, "HOA DON", company_name, "THÁNG 5")
            os.makedirs(month_dir, exist_ok=True)
            
            print(f"Đang tải PDF cho {cont} (Mã PIN: {pin})...")
            pdf_url = f"https://smartport.gemadept.com.vn/apismp/invoice/downloadInvPDF?fkey={pin}&terminal_code={terminal_code}"
            
            try:
                # Dùng cookie của context hiện tại (vì api_context = context.request)
                pdf_resp = await api_context.get(pdf_url, timeout=30000)
                if pdf_resp.status == 200:
                    pdf_bytes = await pdf_resp.body()
                    
                    content_type = pdf_resp.headers.get('content-type', '').lower()
                    if 'application/pdf' in content_type or pdf_bytes[:4] == b'%PDF':
                        safe_cont = "".join(c for c in cont if c.isalnum() or c in ('-', '_')).strip()
                        file_path = os.path.join(month_dir, f"{safe_cont}.pdf")
                        
                        with open(file_path, "wb") as f:
                            f.write(pdf_bytes)
                        print(f"  [v] Đã lưu thành công: {safe_sheet}/{safe_cont}.pdf")
                    else:
                        text_err = pdf_bytes.decode('utf-8', errors='ignore')[:100]
                        print(f"  [!] API không trả về file PDF cho {cont}. Trả về: {text_err}")
                else:
                    print(f"  [!] Lỗi tải PDF cho {cont} (Status: {pdf_resp.status})")
            except Exception as e:
                print(f"  [!] Lỗi quá trình tải PDF cho {cont}: {e}")
                
        print("\n🎉 HOÀN THÀNH!")
        await page.wait_for_timeout(5000)
        await browser.close()

if __name__ == "__main__":
    asyncio.run(main())
