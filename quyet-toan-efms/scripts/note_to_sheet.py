import asyncio
from playwright.async_api import async_playwright

async def main():
    async with async_playwright() as p:
        browser = await p.chromium.launch(headless=False)
        context = await browser.new_context()
        page = await context.new_page()
        
        print("==================================================")
        print("Đang mở Google Sheet...")
        await page.goto("https://docs.google.com/spreadsheets/d/1bBDXbbRRBFrO_C9YTrI3eKUx5aC9Nx01Q9y3IGLstGY/edit#gid=1239449314")
        
        print("\nSẾP VUI LÒNG THỰC HIỆN:")
        print("1. Đăng nhập Google (nếu bị yêu cầu).")
        print("2. Đổi tên sheet thành 'DANH SACH SKILL' (hoặc tạo sheet mới).")
        print("3. CLICK CHUỘT VÀO Ô A1 VÀ ĐỂ YÊN ĐÓ.")
        print("\nRobot sẽ bắt đầu tự động gõ dữ liệu sau 30 GIÂY nữa...")
        print("==================================================")
        
        # Đợi 30 giây cho user đăng nhập và chọn ô
        for i in range(30, 0, -1):
            if i % 5 == 0 or i <= 5:
                print(f"Còn {i} giây...")
            await page.wait_for_timeout(1000)
            
        print("\n[+] BẮT ĐẦU GÕ DỮ LIỆU... VUI LÒNG KHÔNG CHẠM CHUỘT!")
        
        # Gõ Tiêu đề
        await page.keyboard.type("Tên Skill")
        await page.keyboard.press("Tab")
        await page.wait_for_timeout(200)
        await page.keyboard.type("Câu lệnh Kích hoạt")
        await page.keyboard.press("Tab")
        await page.wait_for_timeout(200)
        await page.keyboard.type("Chức năng & Ghi chú")
        await page.keyboard.press("Enter")
        await page.wait_for_timeout(500)
        
        # Gõ Skill 1
        await page.keyboard.type("1. Quyết toán EFMS (Báo cáo phí)")
        await page.keyboard.press("Tab")
        await page.wait_for_timeout(200)
        await page.keyboard.type('uv run --with pandas --with openpyxl --with requests --with xlrd ~/.agent/skills/quyet-toan-efms/scripts/build_efms.py --company "TÊN CÔNG TY" --month 5 --year 2026')
        await page.keyboard.press("Tab")
        await page.wait_for_timeout(200)
        await page.keyboard.type("Tự động quét file Tariff trong thư mục, đối chiếu với API ePort và xuất ra file Excel. Tự phát hiện trùng bill và áp dụng luật riêng cho Haohua Hàng Xuất (bỏ phí BO_OTH_OPS, fix giá 190k).")
        await page.keyboard.press("Enter")
        await page.wait_for_timeout(500)
        
        # Gõ Skill 2
        await page.keyboard.type("2. Tải PDF Hóa Đơn (Smartport)")
        await page.keyboard.press("Tab")
        await page.wait_for_timeout(200)
        await page.keyboard.type("uv run --with playwright --with pandas --with openpyxl --with requests python scripts/download_invoices_api.py")
        await page.keyboard.press("Tab")
        await page.wait_for_timeout(200)
        await page.keyboard.type("Quét TẤT CẢ các sheet ĐANG HIỂN THỊ. Lọc cont Bình Dương, tự động tính mốc thời gian lách luật 3 tháng, tìm mã PIN ngầm và kéo PDF về HOA DON / [Tên Cty] / THÁNG 5. (Cần mồi 1 lần trên web).")
        await page.keyboard.press("Enter")
        
        print("\n🎉 XONG! Đã note dữ liệu thành công! Trình duyệt sẽ đóng sau 10 giây.")
        await page.wait_for_timeout(10000)
        await browser.close()

if __name__ == "__main__":
    asyncio.run(main())
