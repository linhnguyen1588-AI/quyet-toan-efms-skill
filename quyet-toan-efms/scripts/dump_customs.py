import asyncio
from playwright.async_api import async_playwright

async def main():
    async with async_playwright() as p:
        browser = await p.chromium.launch(headless=True)
        page = await browser.new_page()
        print("Visiting page...")
        await page.goto("https://pus.customs.gov.vn/faces/ContainerBarcode", wait_until="networkidle")
        print("Saving HTML...")
        content = await page.content()
        with open("customs_barcode.html", "w", encoding="utf-8") as f:
            f.write(content)
        await browser.close()
        print("Done.")

asyncio.run(main())
