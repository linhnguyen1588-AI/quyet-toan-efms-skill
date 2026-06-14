import asyncio
import sys
if hasattr(sys.stdout, 'reconfigure'):
    sys.stdout.reconfigure(encoding='utf-8')
from playwright.async_api import async_playwright

async def dump():
    async with async_playwright() as p:
        browser = await p.chromium.connect_over_cdp("http://localhost:9222")
        context = browser.contexts[0]
        efms_page = None
        for page in context.pages:
            if "efms" in page.url:
                efms_page = page
                break
        
        if efms_page:
            html = await efms_page.content()
            with open('efms_job_dump.html', 'w', encoding='utf-8') as f:
                f.write(html)
            print("Đã dump HTML ra file efms_job_dump.html")
        else:
            print("Không tìm thấy trang EFMS")

asyncio.run(dump())
