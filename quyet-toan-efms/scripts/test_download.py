import asyncio
import os
from playwright.async_api import async_playwright
import sys

sys.stdout.reconfigure(encoding='utf-8')

async def test_download():
    # Sử dụng tờ khai mà user chắc chắn có dữ liệu
    tk = "308463809530" # Tờ khai này ở task-238 có số 000000000000
    save_dir = r"D:\workspace-ai\CSHT\CSHT HAOHUA\PDFs"
    os.makedirs(save_dir, exist_ok=True)

    async with async_playwright() as p:
        browser = await p.chromium.launch(headless=False)
        context = await browser.new_context(accept_downloads=True)
        page = await context.new_page()

        await page.goto("https://thuphihatang.tphcm.gov.vn/dang-nhap")
        
        try:
            await page.fill('input[type="text"]', "0315428529")
            await page.fill('input[type="password"]', "SSotrans123@1")
        except:
            pass

        print("\n" + "="*70)
        print("VUI LÒNG ĐĂNG NHẬP (Nhập CAPTCHA). Tool sẽ tự chạy sau khi login.")
        print("="*70 + "\n")

        while True:
            try:
                if await page.query_selector('input[name="SO_TK"]'):
                    break
            except:
                pass
            await asyncio.sleep(1)

        print("Đã phát hiện giao diện. Bắt đầu thử nghiệm...")
        
        # Nhập tờ khai và search
        await page.fill('input[name="SO_TK"]', tk)
        await page.keyboard.press("Enter") # Dùng Enter thay vì click cho chắc ăn
        await asyncio.sleep(4)
        
        # Nếu có popup báo lỗi, tắt đi
        try:
            await page.keyboard.press("Escape")
        except:
            pass

        rows = await page.query_selector_all('#TBLDANHSACH tbody tr')
        if not rows:
            print("Không có dòng dữ liệu nào!")
            await asyncio.sleep(10)
            await browser.close()
            return
            
        tds = await rows[0].query_selector_all('td')
        if len(tds) < 11:
            print("Không tìm thấy dữ liệu cho tờ khai này!")
            await asyncio.sleep(10)
            await browser.close()
            return
            
        so_tb = await tds[10].inner_text()
        so_tb = so_tb.strip().split('\n')[0].strip()
        print(f"Số thông báo trước khi bấm: {so_tb}")

        btn_lay = await rows[0].query_selector('.btnXemThongBaoNP')
        if btn_lay:
            print("Bấm Lấy thông báo...")
            await btn_lay.click(force=True)
            await asyncio.sleep(5)
            
            # Tắt popup báo tạo thông báo thành công
            try:
                popups = await page.query_selector_all('.jconfirm-closeIcon')
                for popup in popups:
                    await popup.click()
            except:
                await page.keyboard.press("Escape")

            # Load lại dòng
            rows = await page.query_selector_all('#TBLDANHSACH tbody tr')
            tds = await rows[0].query_selector_all('td')
            so_tb = await tds[10].inner_text()
            so_tb = so_tb.strip().split('\n')[0].strip()
            print(f"Số thông báo sau khi bấm: {so_tb}")

        # Check for Checkbox
        checkbox = await rows[0].query_selector('input[type="checkbox"], input.cb, input.ace, input')
        if checkbox:
            print("Đã tìm thấy Checkbox! Đang tick chọn...")
            await checkbox.check(force=True)
        else:
            print("VẪN KHÔNG CÓ CHECKBOX!")

        print("Bấm nút Tải thông báo (In thông báo NP)...")
        try:
            # Click vào nút In thông báo NP trên toolbar
            dl_btn = await page.query_selector('.btnPrintThongBaoNP')
            if not dl_btn:
                print("Không tìm thấy nút Tải (.btnPrintThongBaoNP)! Thử tìm bằng text...")
                dl_btn = page.locator("text=/.*in thông báo.*/i").first
                
            async with page.expect_download(timeout=10000) as download_info:
                await dl_btn.click(force=True)
            download = await download_info.value
            pdf_path = os.path.join(save_dir, f"{so_tb}.pdf")
            await download.save_as(pdf_path)
            print(f"Tải PDF thành công: {pdf_path}")
        except Exception as e:
            print(f"Lỗi tải PDF: {e}")

        await asyncio.sleep(15)
        await browser.close()

if __name__ == "__main__":
    asyncio.run(test_download())
