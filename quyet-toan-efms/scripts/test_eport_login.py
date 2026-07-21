import asyncio
import os
import sys
from playwright.async_api import async_playwright

sys.stdout.reconfigure(encoding='utf-8')

async def main():
    async with async_playwright() as p:
        browser = await p.chromium.launch(headless=True)
        context = await browser.new_context()
        page = await context.new_page()

        print("[1] Đang mở ePort...")
        await page.goto("https://eport.saigonnewport.com.vn/Home/Login")
        print("URL hiện tại:", page.url)
        print("Tiêu đề:", await page.title())

        # Điền thông tin đăng nhập
        try:
            inputs = await page.locator("input").all()
            print(f"Phát hiện {len(inputs)} ô input")
            for i, inp in enumerate(inputs):
                itype = await inp.get_attribute("type")
                iname = await inp.get_attribute("name")
                iid = await inp.get_attribute("id")
                ipholder = await inp.get_attribute("placeholder")
                print(f"  Input #{i}: id='{iid}', name='{iname}', type='{itype}', placeholder='{ipholder}'")

            await page.fill("input[name='TaxCode']", "0314436809")
            await page.fill("input[name='Password']", "Sotrans1234@")
            
            # Click Login
            await page.click("button[type='submit'], input[type='submit'], .btn-login")
            await page.wait_for_timeout(5000)
            print("URL sau đăng nhập:", page.url)

            # Check search batch API
            api_context = context.request
            resp = await api_context.post("https://eport.saigonnewport.com.vn/BatchNoList/SeachBatchNoList", data={
                "DateFrom": "2026-06-01T00:00:00.000Z",
                "DateTo": "2026-07-21T00:00:00.000Z",
                "OperType": "",
                "HasEDO": "False"
            })
            print("API Status:", resp.status)
            if resp.status == 200:
                print("API Success! Sample response:", (await resp.text())[:300])
            else:
                print("API Failed status:", resp.status)
        except Exception as e:
            print("Lỗi:", e)

        await browser.close()

if __name__ == "__main__":
    asyncio.run(main())
