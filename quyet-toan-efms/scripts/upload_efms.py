import asyncio
import os
import argparse
import sys
if hasattr(sys.stdout, 'reconfigure'):
    sys.stdout.reconfigure(encoding='utf-8')
if hasattr(sys.stderr, 'reconfigure'):
    sys.stderr.reconfigure(encoding='utf-8')
from playwright.async_api import async_playwright

async def upload_file(file_path):
    print(f"Bắt đầu kết nối với trình duyệt qua cổng 9222...")
    async with async_playwright() as p:
        try:
            browser = await p.chromium.connect_over_cdp("http://localhost:9222")
            print("Kết nối thành công!")
        except Exception as e:
            print("LỖI: Không thể kết nối với trình duyệt. Vui lòng đảm bảo trình duyệt đang mở với cờ --remote-debugging-port=9222")
            print(f"Chi tiết lỗi: {e}")
            return
            
        context = browser.contexts[0]
        
        # Tìm tab EFMS
        efms_page = None
        for page in context.pages:
            if "efms.sotrans.com.vn" in page.url:
                efms_page = page
                break
                
        if not efms_page:
            print("LỖI: Không tìm thấy tab nào đang mở trang EFMS (efms.sotrans.com.vn).")
            await browser.close()
            return
            
        print(f"Đã tìm thấy tab EFMS: {efms_page.url}")
        
        # Bring tab to front
        await efms_page.bring_to_front()
        
        print("Đang tìm nút Import và giao diện tải file...")
        # Lấy HTML dump để AI phân tích nếu cần
        html = await efms_page.content()
        with open('efms_dump.html', 'w', encoding='utf-8') as f:
            f.write(html)
        print("Đã lưu giao diện hiện tại vào efms_dump.html")
        
        # Thử một số logic phổ biến:
        # Cách 1: Gán trực tiếp file vào thẻ input[type="file"] đầu tiên tìm thấy
        file_inputs = await efms_page.locator("input[type='file']").all()
        if file_inputs:
            print(f"Tìm thấy {len(file_inputs)} trường tải file (input[type=file]). Đang nhét file vào trường đầu tiên...")
            try:
                await file_inputs[0].set_input_files(file_path)
                print("Đã đính kèm file thành công!")
                
                # Cố gắng tìm nút Confirm/Xác nhận/Import/Upload
                upload_btns = await efms_page.locator("button:has-text('Import'), button:has-text('Upload'), button:has-text('Confirm'), button:has-text('Save'), button:has-text('Đồng ý')").all()
                for btn in upload_btns:
                    if await btn.is_visible():
                        print(f"Phát hiện nút có thể bấm: {await btn.inner_text()} -> Bấm!")
                        await btn.click()
                        break
                        
            except Exception as e:
                print(f"Lỗi khi đính kèm file: {e}")
        else:
            print("Không tìm thấy thẻ input[type='file'] nào ẩn trên trang.")
            print("Sếp có thể cần BẤM MỞ popup Import trước khi chạy kịch bản này!")

        print("Hoàn thành quá trình.")
        await browser.close()

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="EFMS Auto Uploader")
    parser.add_argument("--file", default=r"D:\workspace-ai\QUYET TOAN 2026\LogisticsImport_EFMS_Buying_Haohua Hàng Nhập_05_2026.xlsx", help="Đường dẫn file Excel cần upload")
    args = parser.parse_args()
    
    if not os.path.exists(args.file):
        print(f"CẢNH BÁO: Không tìm thấy file {args.file}")
    
    asyncio.run(upload_file(args.file))
