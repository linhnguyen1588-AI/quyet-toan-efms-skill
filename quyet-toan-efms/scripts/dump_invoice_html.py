import asyncio
from playwright.async_api import async_playwright

async def main():
    async with async_playwright() as p:
        browser = await p.chromium.launch(headless=False)
        context = await browser.new_context()
        page = await context.new_page()
        
        print("Trình duyệt đã mở. Vui lòng đăng nhập và vào trang Hóa đơn điện tử.")
        print("Nhập 1 mã PIN (VD: BDP26030339713), bấm Tìm kiếm để ra kết quả.")
        print("Xong thì gõ 'done' vào cửa sổ terminal hoặc đợi 30 giây để lưu HTML...")
        
        await page.goto("https://smartport.gemadept.com.vn/")
        
        # Đợi 30 giây để user thao tác
        await page.wait_for_timeout(30000)
        
        html = await page.content()
        with open("invoice_page.html", "w", encoding="utf-8") as f:
            f.write(html)
            
        print("Đã lưu invoice_page.html! Sếp có thể đóng trình duyệt.")
        await browser.close()

if __name__ == "__main__":
    asyncio.run(main())
