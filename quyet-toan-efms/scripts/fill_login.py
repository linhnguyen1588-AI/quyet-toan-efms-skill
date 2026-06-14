import asyncio
import sys
if hasattr(sys.stdout, 'reconfigure'):
    sys.stdout.reconfigure(encoding='utf-8')
if hasattr(sys.stderr, 'reconfigure'):
    sys.stderr.reconfigure(encoding='utf-8')
from playwright.async_api import async_playwright

async def fill_login():
    try:
        async with async_playwright() as p:
            # Connect to existing browser
            browser = await p.chromium.connect_over_cdp("http://localhost:9222")
            context = browser.contexts[0]
            
            # Find the EFMS login page
            login_page = None
            for page in context.pages:
                if "efms.sotrans.com.vn" in page.url and "login" in page.url.lower():
                    login_page = page
                    break
                    
            if not login_page:
                print("Lỗi: Không tìm thấy trang đăng nhập EFMS đang mở.")
                return
                
            print(f"Đã kết nối thành công tới tab: {login_page.url}")
            await login_page.bring_to_front()
            
            # Try to find input fields. Usually there is a username and a password field.
            print("Đang điền Username...")
            await login_page.fill("input[name='username'], input[placeholder*='Username'], input[type='text']", "leo.linh", timeout=5000)
            
            print("Đang điền Password...")
            await login_page.fill("input[name='password'], input[placeholder*='Password'], input[type='password']", "12345678", timeout=5000)
            
            print("Đã điền xong! Đang chờ bạn tự bấm Đăng nhập...")
            
    except Exception as e:
        print(f"Lỗi: {e}")

asyncio.run(fill_login())
