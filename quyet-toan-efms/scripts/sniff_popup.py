import asyncio
from playwright.async_api import async_playwright
import json
import os

async def main():
    async with async_playwright() as p:
        browser = await p.chromium.launch(headless=False)
        context = await browser.new_context()
        page = await context.new_page()
        
        print("==================================================")
        print("Trình duyệt đã mở. Sếp vui lòng làm lại các bước:")
        print("1. Đăng nhập -> Vào Hóa đơn điện tử.")
        print("2. Tìm 1 mã PIN.")
        print("3. Bấm Tải PDF (sẽ mở tab mới).")
        print("Robot sẽ chụp lại link của tab mới đó!")
        print("==================================================")
        
        api_data = {}
        
        # Bắt sự kiện mở tab mới
        async def on_page(new_page):
            await new_page.wait_for_load_state()
            url = new_page.url
            print(f"\n[+] ĐÃ TÓM ĐƯỢC LINK TAB MỚI: {url}")
            api_data["url"] = url
            with open("download_link.json", "w", encoding="utf-8") as f:
                json.dump(api_data, f, indent=2, ensure_ascii=False)
                
        context.on("page", on_page)
        
        await page.goto("https://smartport.gemadept.com.vn/")
        
        # Đợi đến khi file được tạo
        for _ in range(120):
            if os.path.exists("download_link.json"):
                print("Đã lấy được link! Chuẩn bị đóng trình duyệt...")
                await asyncio.sleep(2)
                break
            await asyncio.sleep(1)
            
        await browser.close()

if __name__ == "__main__":
    if os.path.exists("download_link.json"):
        os.remove("download_link.json")
    asyncio.run(main())
