import asyncio
from playwright.async_api import async_playwright
import json

async def main():
    async with async_playwright() as p:
        browser = await p.chromium.launch(headless=False)
        context = await browser.new_context()
        page = await context.new_page()
        
        print("Đang mở trình duyệt để bắt API...")
        
        async def handle_request(request):
            if "api/eir/" in request.url or "api/invoice/" in request.url:
                try:
                    post_data = request.post_data
                    with open("api_sniff.log", "a", encoding="utf-8") as f:
                        f.write(f"[{request.method}] {request.url}\n")
                        if post_data:
                            f.write(f"Payload: {post_data}\n")
                except:
                    pass

        async def handle_response(response):
            if "api/eir/" in response.url or "api/invoice/" in response.url:
                try:
                    text = await response.text()
                    with open("api_sniff.log", "a", encoding="utf-8") as f:
                        f.write(f"[RESPONSE] {response.url}\n")
                        f.write(f"Body: {text[:1000]}\n")
                except:
                    pass
                    
        page.on("request", handle_request)
        page.on("response", handle_response)
        
        await page.goto("https://smartport.gemadept.com.vn/")
        
        # Giữ browser mở để tương tác
        while True:
            await asyncio.sleep(1)

if __name__ == "__main__":
    asyncio.run(main())
