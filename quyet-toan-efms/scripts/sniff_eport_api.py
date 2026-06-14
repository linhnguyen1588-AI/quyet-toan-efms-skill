import asyncio
from playwright.async_api import async_playwright
import json
import os
import sys

# Force UTF-8 encoding for Windows console to prevent print() Unicode errors
sys.stdout.reconfigure(encoding='utf-8')

async def main():
    # Tạo thư mục lưu video nếu chưa có
    if not os.path.exists("video_recordings"):
        os.makedirs("video_recordings")
        
    async with async_playwright() as p:
        # Bật trình duyệt có UI (headless=False)
        browser = await p.chromium.launch(headless=False)
        
        # Bật tính năng quay video màn hình lưu vào thư mục 'video_recordings'
        context = await browser.new_context(
            record_video_dir="video_recordings",
            record_video_size={"width": 1280, "height": 720}
        )
        page = await context.new_page()
        
        print("==================================================")
        print("Trình duyệt đã mở và ĐANG QUAY PHIM MÀN HÌNH.")
        print("Sếp vui lòng làm các bước sau:")
        print("1. Đăng nhập ePort.")
        print("2. Vào trang Hóa đơn điện tử.")
        print("3. Nhập dữ liệu tìm kiếm (Số Hóa Đơn hoặc Số Đăng Ký) rồi bấm Tìm Kiếm.")
        print("4. BẤM TẢI PDF HÓA ĐƠN VỀ MÁY.")
        print("Robot đang rình sẵn để copy link tải bí mật và quay lại toàn bộ thao tác...")
        print("==================================================")
        
        api_data = {}
        
        async def on_request(request):
            url = request.url.lower()
            # Bắt các API liên quan đến tải hóa đơn, export, view pdf...
            if "invoice" in url or "hoadon" in url or "download" in url or "export" in url or "print" in url:
                print(f"[Network] Đang tải: {request.url}")
                if "pdf" in url or request.method == "POST":
                    print(f"\n[+] ĐÃ TÓM ĐƯỢC API TẢI HÓA ĐƠN: {request.url}")
                    api_data["url"] = request.url
                    api_data["method"] = request.method
                    api_data["headers"] = request.headers
                    api_data["post_data"] = request.post_data
                    
                    with open("eport_download_api.json", "w", encoding="utf-8") as f:
                        json.dump(api_data, f, indent=2, ensure_ascii=False)
                
        page.on("request", on_request)
        
        await page.goto("https://eport.saigonnewport.com.vn/")
        
        # Đợi đến khi sếp thao tác xong (Tối đa 5 phút)
        for _ in range(300):
            if os.path.exists("eport_download_api.json"):
                print("\n[OK] Đã lấy được bí kíp API tải hóa đơn! Trình duyệt sẽ tự đóng sau 5 giây để lưu Video...")
                await asyncio.sleep(5)
                break
            await asyncio.sleep(1)
            
        await context.close()  # Phải đóng context trước thì video mới được lưu file hoàn chỉnh
        await browser.close()
        
        print("\n🎉 XONG! File Video đã được lưu trong thư mục 'video_recordings'.")

if __name__ == "__main__":
    if os.path.exists("eport_download_api.json"):
        os.remove("eport_download_api.json")
    asyncio.run(main())
