import asyncio
import json
import argparse
import os
import sys
from playwright.async_api import async_playwright

sys.stdout.reconfigure(encoding='utf-8')

api_logs = []

async def handle_response(response):
    if "google-analytics" in response.url or response.url.endswith(('.js', '.css', '.png', '.jpg', '.woff')):
        return
    try:
        req = response.request
        log = {
            "url": req.url,
            "method": req.method,
            "headers": req.headers,
            "post_data": req.post_data,
            "status": response.status
        }
        try:
            content_type = response.headers.get("content-type", "")
            if "application/json" in content_type:
                log["response_body"] = await response.json()
            elif "text/" in content_type:
                text = await response.text()
                if len(text) < 10000:  # Tránh lưu nguyên cái trang HTML bự chảng
                    log["response_body"] = text
        except:
            pass
        api_logs.append(log)
    except Exception as e:
        pass

async def handle_download(download):
    api_logs.append({"url": download.url, "method": "DOWNLOAD EVENT"})

async def main():
    async with async_playwright() as p:
        browser = await p.chromium.launch(headless=False)
        context = await browser.new_context(accept_downloads=True)
        page = await context.new_page()
        
        page.on("response", handle_response)
        page.on("download", handle_download)
        
        print("\n==================================================")
        print("MÁY QUAY LÉN ĐÃ BẬT!")
        print("Sếp thao tác Tìm kiếm & Tải hóa đơn cho Book thứ 2: 258527559")
        print("Xong sếp chat 'xong', em sẽ tự dừng và phân tích nhé.")
        print("==================================================\n")
        
        await page.goto("https://eport.saigonnewport.com.vn/FullContainerDelivery")
        
        if os.path.exists('stop.txt'):
            os.remove('stop.txt')
            
        while not os.path.exists('stop.txt'):
            await page.wait_for_timeout(1000)
            
        with open("eport_intercept_logs.json", "w", encoding="utf-8") as f:
            json.dump(api_logs, f, indent=4, ensure_ascii=False)
            
        print("Đã nhận được tín hiệu dừng. Đã lưu lịch sử API vào eport_intercept_logs.json.")
        await browser.close()

if __name__ == "__main__":
    asyncio.run(main())
