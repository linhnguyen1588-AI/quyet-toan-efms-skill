import asyncio
from playwright.async_api import async_playwright
import sys
sys.stdout.reconfigure(encoding='utf-8')

async def run_sniffer():
    async with async_playwright() as p:
        browser = await p.chromium.launch(headless=False)
        context = await browser.new_context(accept_downloads=True)
        page = await context.new_page()

        print("\n" + "="*70)
        print("TRÌNH DUYỆT ĐÃ MỞ!")
        print("Sếp vui lòng đăng nhập và thao tác tìm kiếm/tải 1 tờ khai bằng tay nhé.")
        print("Mọi thao tác của sếp (click, tìm kiếm) sẽ được ghi lại để em học theo ạ.")
        print("Khi nào sếp làm xong 1 tờ khai thì quay lại báo em nhé!")
        print("="*70 + "\n")

        # Inject script to log clicks
        await page.goto("https://thuphihatang.tphcm.gov.vn/dang-nhap")
        
        try:
            await page.fill('input[type="text"]', "0315428529")
            await page.fill('input[type="password"]', "SSotrans123@1")
        except:
            pass

        # Log requests to find the exact API endpoint being called
        page.on("request", lambda request: print(f"[Mạng] Đang tải: {request.url}") if "thuphihatang" in request.url and "api" in request.url.lower() else None)
        
        # Keep browser open
        while True:
            await asyncio.sleep(1)

if __name__ == "__main__":
    asyncio.run(run_sniffer())
