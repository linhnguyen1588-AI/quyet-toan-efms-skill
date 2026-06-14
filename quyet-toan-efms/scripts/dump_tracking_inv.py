import asyncio
from playwright.async_api import async_playwright

async def main():
    async with async_playwright() as p:
        browser = await p.chromium.launch(headless=False)
        context = await browser.new_context()
        page = await context.new_page()
        
        await page.goto("https://smartport.gemadept.com.vn/login")
        print("Trình duyệt đã mở. Sếp vui lòng làm các bước sau:")
        print("1. Đăng nhập (với Captcha).")
        print("2. Tự bấm sang trang 'Tra cứu số hóa đơn' (tracking_inv).")
        print("3. Điền thử 1 mã PinCode (VD: BDP26041323643) và bấm Tìm kiếm.")
        print("4. Khi màn hình hiện ra kết quả hóa đơn (có nút Tải), sếp hãy để nguyên màn hình đó.")
        print("Script sẽ tự động chụp HTML sau 30 giây nữa...")
        
        await page.wait_for_timeout(30000)
        
        html = await page.content()
        with open('tracking_inv_dump.html', 'w', encoding='utf-8') as f:
            f.write(html)
            
        print("Đã lưu tracking_inv_dump.html thành công! Trình duyệt sẽ đóng.")
        await browser.close()

if __name__ == "__main__":
    asyncio.run(main())
