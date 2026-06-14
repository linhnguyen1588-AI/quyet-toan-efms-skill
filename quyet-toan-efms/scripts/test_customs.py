import asyncio
from playwright.async_api import async_playwright

async def main():
    async with async_playwright() as p:
        browser = await p.chromium.launch(headless=True)
        page = await browser.new_page()
        print("Visiting page...")
        await page.goto("https://pus.customs.gov.vn/faces/ContainerBarcode", wait_until="networkidle")
        
        # Fill data
        await page.fill('input[name="pt1:it1"]', "0305623305")
        await page.fill('input[name="pt1:it2"]', "123456789012")
        await page.fill('input[name="pt1:it3"]', "02CI")
        # Date field has id pt1:it4::content
        await page.fill('input[name="pt1:it4"]', "02/06/2023")
        
        # Click search
        print("Clicking search...")
        await page.click('a.xfp') # "Lấy thông tin" button
        
        # Wait for either table data or a popup
        await asyncio.sleep(5)
        
        content = await page.content()
        with open("test_result.html", "w", encoding="utf-8") as f:
            f.write(content)
        
        # Check for alert
        alert = await page.query_selector('.x1g8.af_messages_dialog')
        if alert and await alert.is_visible():
            text = await alert.inner_text()
            print("Alert found:", text)
        else:
            print("No alert popup.")
            
        await browser.close()

asyncio.run(main())
