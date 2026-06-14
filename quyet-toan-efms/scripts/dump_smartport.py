import asyncio
from playwright.async_api import async_playwright
import os

async def dump_html():
    async with async_playwright() as p:
        # Launch browser in headless mode
        browser = await p.chromium.launch(headless=True)
        page = await browser.new_page()
        
        print("Navigating to Smartport...")
        await page.goto("https://smartport.gemadept.com.vn/tracking/eirsrv", wait_until="networkidle")
        
        # Wait a bit more for React/Angular to render fully
        await page.wait_for_timeout(3000)
        
        # Dump the HTML
        html_content = await page.content()
        
        output_file = os.path.join(os.path.dirname(__file__), "smartport_dump.html")
        with open(output_file, "w", encoding="utf-8") as f:
            f.write(html_content)
            
        print(f"HTML dumped to {output_file}")
        
        await browser.close()

if __name__ == "__main__":
    asyncio.run(dump_html())
