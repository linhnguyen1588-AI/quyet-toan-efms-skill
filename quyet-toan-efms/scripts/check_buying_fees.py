import asyncio
import os
import sys
if hasattr(sys.stdout, 'reconfigure'):
    sys.stdout.reconfigure(encoding='utf-8')
if hasattr(sys.stderr, 'reconfigure'):
    sys.stderr.reconfigure(encoding='utf-8')

import pandas as pd
from playwright.async_api import async_playwright

async def run_checker():
    file_path = r"D:\Standard ReportVND-100626.xlsx"
    if not os.path.exists(file_path):
        print(f"Lỗi: Không tìm thấy file {file_path}")
        return
        
    print(f"Đang đọc file {file_path}...")
    df = pd.read_excel(file_path, header=None)
    
    # Header ở dòng 0, JOB ID ở cột 1
    job_ids = []
    for r in range(1, len(df)):
        val = str(df.iloc[r, 1]).strip()
        if val and val.lower() != 'nan' and len(val) > 5:
            job_ids.append(val)
            
    print(f"Tìm thấy {len(job_ids)} JOBs cần kiểm tra.")
    
    missing_jobs = []
    
    async with async_playwright() as p:
        try:
            browser = await p.chromium.connect_over_cdp("http://localhost:9222")
            context = browser.contexts[0]
            
            # Find EFMS page
            efms_page = None
            for page in context.pages:
                if "efms.sotrans.com.vn" in page.url:
                    efms_page = page
                    break
                    
            if not efms_page:
                print("Lỗi: Không tìm thấy tab nào mở EFMS. Vui lòng mở EFMS trước.")
                return
                
            await efms_page.bring_to_front()
            
            for i, job_id in enumerate(job_ids):
                print(f"[{i+1}/{len(job_ids)}] Đang kiểm tra JOB: {job_id}")
                
                # Điều hướng tới trang chi tiết Job (sử dụng custom-clearance/detail)
                detail_url = f"https://efms.sotrans.com.vn/en/#/home/operation/custom-clearance/detail?jobNo={job_id}"
                await efms_page.goto(detail_url, wait_until="networkidle")
                
                # Đợi trang tải xong, cố gắng tìm thẻ "Buying"
                try:
                    await efms_page.wait_for_selector("text='Buying', a:has-text('Buying'), span:has-text('Buying')", timeout=10000)
                    
                    # Bấm vào tab Buying
                    elements = await efms_page.locator("text='Buying', a:has-text('Buying'), span:has-text('Buying')").all()
                    for el in elements:
                        if await el.is_visible():
                            await el.click()
                            break
                            
                    # Chờ dữ liệu bảng Buying tải (khoảng 3 giây)
                    await asyncio.sleep(3)
                    
                    html = await efms_page.content()
                    
                    has_nof = "BO_NOF_OPS" in html
                    has_nde = "BO_NDE_OPS" in html
                    
                    if not has_nof or not has_nde:
                        missing = []
                        if not has_nof: missing.append("BO_NOF_OPS")
                        if not has_nde: missing.append("BO_NDE_OPS")
                        print(f" => THIẾU PHÍ: {', '.join(missing)}")
                        missing_jobs.append({
                            "Job": job_id,
                            "Thiếu": ", ".join(missing)
                        })
                    else:
                        print(" => Đã đủ 2 phí.")
                        
                except Exception as e:
                    print(f" => Lỗi không thể truy cập hoặc tìm thấy tab Buying: {e}")
                    missing_jobs.append({
                        "Job": job_id,
                        "Thiếu": "Lỗi truy cập"
                    })
                    
            print("\n" + "="*50)
            print("KẾT QUẢ KIỂM TRA")
            print("="*50)
            if not missing_jobs:
                print("Tất cả các JOB đều đã có đủ phí BO_NOF_OPS và BO_NDE_OPS!")
            else:
                for m in missing_jobs:
                    print(f"JOB {m['Job']} - Thiếu: {m['Thiếu']}")
                    
            # Export to Excel
            if missing_jobs:
                df_out = pd.DataFrame(missing_jobs)
                out_path = r"D:\JOB_Thieu_Phi.xlsx"
                df_out.to_excel(out_path, index=False)
                print(f"Đã xuất danh sách JOB thiếu phí ra file: {out_path}")
                
        except Exception as e:
            print(f"Lỗi khi thực thi: {e}")

if __name__ == "__main__":
    asyncio.run(run_checker())
