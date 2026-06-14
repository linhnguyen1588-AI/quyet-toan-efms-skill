# Quy Trình Quyết Toán EFMS (Đặc thù công ty PARITAS và TOÀN LỰC)

Tài liệu này ghi chú lại cấu hình và các quy tắc đặc thù (Business Logic) khi chạy lệnh xuất file Excel báo cáo quyết toán cho phần mềm EFMS đối với hai khách hàng: **PARITAS** và **TOÀN LỰC**.

## 1. Mục đích và Công cụ
- **File script:** `scripts/build_efms.py`
- **Chức năng:** Lấy dữ liệu từ file Excel (Google Sheet tải về) hoặc trực tiếp từ Google Sheets API, và đối chiếu với bảng giá Tariff để tự động tính phí, sau đó xuất ra file chuẩn của hệ thống EFMS.
- **Lệnh thực thi chuẩn:** 
  ```bash
  uv run --with pandas --with openpyxl --with requests --with python-calamine scripts/build_efms.py --company "Toàn Lực" --month 5 --year 2026
  ```

## 2. Quy tắc Tính phí và Luồng Mặc Định
Đối với PARITAS và TOÀN LỰC, có một số quy tắc ngầm định quan trọng:
- **Luồng mặc định:** Mặc định phân luồng hải quan cho mọi tờ khai của Paritas và Toàn Lực là **Luồng Vàng**.
- **Bảng giá Khoán (BO_NOF_OPS):**
  - 1 Container: 450,000 VNĐ
  - 2 Container: 400,000 VNĐ
  - 3 Container: 380,000 VNĐ
  - 4 Container: 380,000 VNĐ
  - Từ 5 Container trở lên: 370,000 VNĐ

## 3. Quy tắc Lấy dữ liệu từ Google Sheets API
- **Lọc rác:** Tool sẽ tự động bỏ qua (skip) các dòng không có số Tờ Khai (cột `to_khai` rỗng hoặc `NaN`).
- **Phụ thu Hải quan và Bốc xếp:**
  - **Cột O (Phụ thu hải quan):** Được API trả về dưới tên key `dest_hm_20`. Tool sẽ tự động lấy và gán thành charge code `BO_OTH_OPS` với ghi chú "phụ thu hải quan".
  - **Cột P (Phụ thu bốc xếp):** Được API trả về dưới tên key `dest_hm_40`. Tool sẽ tự động lấy và gán thành charge code `BO_OTH_OPS` với ghi chú "Phụ thu bốc xếp".

## 4. Quy tắc Tính số lượng Container
- Hệ thống sẽ lấy chính xác số lượng container thực tế bằng cách:
  - Cột F: Số lượng Container 20' (20'DC)
  - Cột G: Số lượng Container 40' (40'DC)
- **Tổng số Cont = Cột F + Cột G.** (Hệ thống đã fix lỗi lấy nhầm sang Cột J - Phụ thu dầu).

## 5. Quy tắc Sắp xếp và Đánh dấu màu (Highlight)
File Excel xuất ra (`LogisticsImport_EFMS_Buying_<Tên>_*.xlsx`) được tự động hóa trình bày để kế toán dễ nhìn nhất:

1. **Phân loại Ghi chú (Note):**
   - Các dòng cước phí CÓ GHI CHÚ (cột Note có dữ liệu, ví dụ "Phụ thu bốc xếp", "Phụ thu hải quan") sẽ bị đẩy xuống **dưới cùng** của file.
   - Các dòng cước phí KHÔNG CÓ GHI CHÚ sẽ nằm ở phần trên.

2. **Gom nhóm theo Charge Code:**
   Bên trong từng phần (có note và không có note), các lô hàng được gom lại theo từng Charge Code:
   - Nhóm 1: `BO_COF_OPS` (Phí tờ khai)
   - Nhóm 2: `BO_NOF_OPS` (Phí làm hàng/cước cont)
   - Nhóm 3: `BO_OTH_OPS` (Phụ thu khác)

3. **Tô màu (Highlight):**
   - `BO_COF_OPS`: Nền **xanh dương nhạt**.
   - `BO_NOF_OPS`: Nền **xanh lá cây nhạt**.
   - `BO_OTH_OPS`: Nền **vàng/cam nhạt**.
   - **Trùng Bill (Duplicate HBL):** Bôi **nền vàng chói** ngay tại cột `Số Bill` (Cột A) để cảnh báo lô hàng bị trùng lặp.
