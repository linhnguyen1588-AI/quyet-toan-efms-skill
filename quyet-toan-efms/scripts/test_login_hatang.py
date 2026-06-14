import asyncio
from playwright.async_api import async_playwright

async def main():
    async with async_playwright() as p:
        browser = await p.chromium.launch(headless=True)
        context = await browser.new_context()
        page = await context.new_page()
        
        print("Navigating to login page...")
        await page.goto("https://thuphihatang.tphcm.gov.vn/dang-nhap")
        await page.wait_for_load_state("networkidle")
        
        # Save HTML to inspect login form fields
        html = await page.content()
        with open("hatang_login_page.html", "w", encoding="utf-8") as f:
            f.write(html)
        print("Saved login page HTML to hatang_login_page.html")
        
        # Try to find common login fields
        # username: 0315428529
        # pass: SSotrans123@1
        
        # Fill username
        try:
            # Often it's an input with type="text" or name="username" or id="username"
            await page.fill('input[type="text"]', "0315428529")
            await page.fill('input[type="password"]', "SSotrans123@1")
            print("Filled credentials")
            
            # Click login button
            await page.click('button[type="submit"], button:has-text("Đăng nhập"), button:has-text("Login")')
            print("Clicked login")
            
            await page.wait_for_timeout(5000) # wait for navigation or error
            
            html_after = await page.content()
            with open("hatang_after_login.html", "w", encoding="utf-8") as f:
                f.write(html_after)
            print("Saved after login HTML to hatang_after_login.html")
            
        except Exception as e:
            print("Error during login:", e)
        
        await browser.close()

if __name__ == "__main__":
    asyncio.run(main())
