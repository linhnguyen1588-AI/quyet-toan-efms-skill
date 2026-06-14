import asyncio
import json
import argparse
from playwright.async_api import async_playwright

api_logs = []

async def handle_response(response):
    if "google-analytics" in response.url or "socket.io" in response.url or response.url.endswith('.js') or response.url.endswith('.css'):
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
            is_json = "application/json" in content_type
            if is_json:
                log["response_body"] = await response.json()
        except:
            pass
        api_logs.append(log)
        print(f"Bắt được URL: {req.method} {req.url}")
    except Exception as e:
        pass

async def handle_download(download):
    print(f"BẮT ĐƯỢC SỰ KIỆN DOWNLOAD! URL: {download.url}")
    api_logs.append({"url": download.url, "method": "DOWNLOAD EVENT"})

async def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--user", type=str, default="")
    parser.add_argument("--password", type=str, default="")
    args = parser.parse_args()

    async with async_playwright() as p:
        browser = await p.chromium.launch(headless=False)
        context = await browser.new_context(accept_downloads=True)
        page = await context.new_page()
        
        page.on("response", handle_response)
        page.on("download", handle_download)
        
        await page.goto("https://smartport.gemadept.com.vn/")
        
        if args.user and args.password:
            try:
                await page.fill("input[placeholder*='Tài khoản'], input[name='username']", args.user, timeout=5000)
                await page.fill("input[placeholder*='Mật khẩu'], input[type='password']", args.password, timeout=5000)
            except:
                pass
        
        print("Script đang ghi hình ngầm. Hãy tiến hành tra cứu và tải hóa đơn trên trình duyệt.")
        print("Khi nào xong, sếp chat lại với em, em sẽ tự động dừng script và phân tích nhé.")
        
        import os
        if os.path.exists('stop.txt'):
            os.remove('stop.txt')
            
        while not os.path.exists('stop.txt'):
            await page.wait_for_timeout(1000)
            
        with open("api_intercept_logs.json", "w", encoding="utf-8") as f:
            json.dump(api_logs, f, indent=4, ensure_ascii=False)
            
        print("Đã nhận được tín hiệu dừng. Đã lưu lịch sử API vào api_intercept_logs.json.")
        await browser.close()

if __name__ == "__main__":
    asyncio.run(main())
