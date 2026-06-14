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
        print("Trình duyệt đã mở. Sếp vui lòng làm các bước sau:")
        print("1. Đăng nhập.")
        print("2. Vào trang Hóa đơn điện tử.")
        print("3. Tự nhập 1 mã PIN (VD: BDP26030339713), bấm Tìm.")
        print("4. BẤM TAY VÀO NÚT TẢI PDF (hoặc nút Xem).")
        print("Robot đang rình sẵn để copy link tải bí mật...")
        print("==================================================")
        
        api_data = {}
        
        async def on_request(request):
            url = request.url.lower()
            if "invoice" in url and ("pdf" in url or "download" in url or "export" in url or "print" in url):
                print(f"\n[+] ĐÃ TÓM ĐƯỢC API TẢI HÓA ĐƠN: {request.url}")
                api_data["url"] = request.url
                api_data["method"] = request.method
                api_data["headers"] = request.headers
                api_data["post_data"] = request.post_data
                with open("download_api.json", "w", encoding="utf-8") as f:
                    json.dump(api_data, f, indent=2, ensure_ascii=False)
                
        page.on("request", on_request)
        page.on("response", lambda r: None) # just to ensure network events are tracked
        
        await page.goto("https://smartport.gemadept.com.vn/")
        
        # Đợi đến khi file được tạo (tối đa 120 giây)
        for _ in range(120):
            if os.path.exists("download_api.json"):
                print("Đã lấy được bí kíp! Chuẩn bị đóng trình duyệt...")
                await asyncio.sleep(2)
                break
            await asyncio.sleep(1)
            
        await browser.close()

if __name__ == "__main__":
    if os.path.exists("download_api.json"):
        os.remove("download_api.json")
    asyncio.run(main())
