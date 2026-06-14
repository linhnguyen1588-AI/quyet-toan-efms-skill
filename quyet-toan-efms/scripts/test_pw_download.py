import asyncio
from playwright.async_api import async_playwright

async def main():
    async with async_playwright() as p:
        browser = await p.chromium.launch(headless=False)
        context = await browser.new_context(accept_downloads=True)
        page = await context.new_page()
        
        await page.goto("https://smartport.gemadept.com.vn/")
        print("Vui lòng nhập Captcha và Đăng nhập...")
        
        # Chờ user đăng nhập (chờ URL đổi hoặc user báo xong)
        try:
            await page.wait_for_function("window.location.pathname !== '/login' && window.location.pathname !== '/'", timeout=45000)
            print("Đăng nhập thành công!")
        except:
            print("Chưa thấy chuyển trang, tiếp tục...")
            
        print("Đang truy cập trang tracking_inv...")
        await page.goto("https://smartport.gemadept.com.vn/tracking/tracking_inv")
        await page.wait_for_timeout(3000)
        
        print("Điền mã PinCode BDP26041323643...")
        try:
            await page.locator("#tracking").fill("BDP26041323643")
            await page.locator("button.ant-input-search-button").click()
            print("Đã bấm Tìm kiếm, đợi kết quả...")
            
            # Đợi icon download xuất hiện (tối đa 15s)
            await page.locator(".anticon-download").wait_for(timeout=15000)
            print("Đã thấy nút Tải Hóa Đơn! Bắt đầu bấm tải...")
            
            async with page.expect_download() as download_info:
                await page.locator(".anticon-download").first.click()
            download = await download_info.value
            
            path = await download.path()
            print(f"THÀNH CÔNG! Đã tải file: {download.suggested_filename} lưu tại {path}")
            print(f"URL tải file: {download.url}")
            
        except Exception as e:
            print("LỖI HOẶC YÊU CẦU CAPTCHA:", e)
            print("Đang chụp ảnh màn hình để debug...")
            await page.screenshot(path="debug_error.png")
            
        await page.wait_for_timeout(5000)
        await browser.close()

if __name__ == "__main__":
    asyncio.run(main())
