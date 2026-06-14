import asyncio
import os
import argparse
import pandas as pd
from playwright.async_api import async_playwright

async def download_invoices(sheet_url, output_dir, username, password):
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)
    
    # 1. Đọc dữ liệu từ Google Sheets
    print(f"Đang đọc dữ liệu từ Google Sheets: {sheet_url}")
    try:
        # Chuyển đổi URL sang dạng export CSV
        csv_url = sheet_url.replace('/edit?', '/export?format=csv&').replace('#gid=', '&gid=')
        if '/edit#gid=' in sheet_url:
            csv_url = sheet_url.replace('/edit#gid=', '/export?format=csv&gid=')
            
        df = pd.read_csv(csv_url)
        # Giả sử cột chứa số cont có tên là 'Container' hoặc lấy cột đầu tiên
        # Tạm thời in ra các cột để kiểm tra
        print("Các cột trong file:", df.columns.tolist())
        # Tạm lấy cột đầu tiên nếu không có cột 'Container'
        cont_col = 'SỐ CONT' if 'SỐ CONT' in df.columns else df.columns[0]
        containers = df[cont_col].dropna().astype(str).tolist()
        print(f"Đã lấy được {len(containers)} số container: {containers[:5]}...")
    except Exception as e:
        print(f"Lỗi đọc Google Sheets: {e}")
        return

    # 2. Mở trình duyệt và đăng nhập
    async with async_playwright() as p:
        browser = await p.chromium.launch(headless=False)
        context = await browser.new_context()
        page = await context.new_page()
        
        print("\nĐang truy cập Smartport...")
        await page.goto("https://smartport.gemadept.com.vn/", wait_until="networkidle")
        
        # Điền thông tin đăng nhập
        print("Đang điền thông tin đăng nhập...")
        try:
            # Các selector này có thể cần điều chỉnh sau khi xem xét DOM thực tế
            await page.fill("input[placeholder*='Tài khoản'], input[name='username'], input[id='username']", username, timeout=5000)
            await page.fill("input[placeholder*='Mật khẩu'], input[type='password']", password, timeout=5000)
            print("Đã điền Username và Password.")
            print(">>> VUI LÒNG NHẬP CAPTCHA TRÊN TRÌNH DUYỆT VÀ BẤM ĐĂNG NHẬP <<<")
        except Exception as e:
            print("Không tìm thấy ô đăng nhập tự động, vui lòng tự nhập bằng tay.")
        
        # Đợi người dùng tự đăng nhập và trang load xong (Ví dụ đợi nút Đăng xuất xuất hiện hoặc URL thay đổi)
        # Ở đây ta sẽ đợi đến khi URL chứa '/dashboard' hoặc tương tự, hoặc đơn giản là cho người dùng thời gian 60s
        print("Đang chờ đăng nhập thành công (Thời gian chờ 60s)...")
        try:
            await page.wait_for_function("window.location.pathname !== '/login' && window.location.pathname !== '/'", timeout=60000)
            print("Đăng nhập thành công!")
        except Exception as e:
            print("Hết thời gian chờ đăng nhập, có thể người dùng đã đóng trình duyệt, hoặc không phát hiện được trạng thái.")
            
        print("Đang chuyển hướng tới trang Tra cứu EIR/Hóa đơn...")
        await page.goto("https://smartport.gemadept.com.vn/tracking/eirsrv", wait_until="networkidle")
        
        print("Đang chờ trang tải nội dung (10s)...")
        await page.wait_for_timeout(10000)
        
        # Lưu lại HTML lần cuối cùng để chắc chắn lấy được DOM của trang hóa đơn
        try:
            html = await page.content()
            with open('eirsrv_dump.html', 'w', encoding='utf-8') as f:
                f.write(html)
            print("Đã lưu HTML của trang Hóa đơn (eirsrv_dump.html).")
        except Exception:
            pass
        
        # Đóng trình duyệt tạm thời để agent phân tích
        await browser.close()
        
        # 3. TODO: Điều hướng đến trang Tra cứu hóa đơn / EIR và tải file
        # Sẽ cần phân tích DOM của trang Tra cứu sau khi login
        
        print("Hoàn thành giai đoạn 1. Đang đóng trình duyệt...")
        await browser.close()

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Tải hóa đơn điện tử từ Smartport")
    parser.add_argument("--sheet", required=True, help="URL của Google Sheets")
    parser.add_argument("--output", default="D:\\Danh_sach_hoa_don_da_tai", help="Thư mục lưu hóa đơn")
    parser.add_argument("--user", required=True, help="Tài khoản Smartport")
    parser.add_argument("--password", required=True, help="Mật khẩu Smartport")
    args = parser.parse_args()
    
    asyncio.run(download_invoices(args.sheet, args.output, args.user, args.password))
