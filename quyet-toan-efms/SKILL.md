---
name: quyết toán efms
description: Skill dùng để tự động tạo file LogisticsImport EFMS (Buying) dựa trên API Google Sheets và file báo giá Tariff. Hỗ trợ tự động săn tìm (Smart Scan) file Tariff của công ty nếu chúng được lưu chung trong một thư mục lớn.
---

# Hướng Dẫn Kích Hoạt Skill Quyết Toán EFMS (Dành cho Môi trường Production)

Khi người dùng yêu cầu tạo file **LogisticsImport EFMS (Buying)**, hãy thực hiện theo quy trình sau:

## 1. Setup Lần Đầu (Hoặc khi thiếu cấu hình)
Nếu là lần đầu chạy ở một máy mới, bạn hãy kiểm tra nội dung file `config.json` (nằm cùng thư mục với `SKILL.md`). Nếu các trường `tariff_dir` hoặc `output_dir` đang trống, hãy hỏi người dùng:
1. **Thư mục chứa toàn bộ các file báo giá Tariff của khách hàng nằm ở đâu?**
2. **Thư mục lưu trữ các file Excel báo cáo sinh ra nằm ở đâu?**
*(Lưu ý: Nếu người dùng cung cấp đường dẫn tương đối, hãy quy đổi ra đường dẫn tuyệt đối giúp họ).*
Sau khi có câu trả lời, hãy tự cập nhật `config.json` ngay lập tức.

## 2. Quy trình Thực Thi Chính (Tự động hóa hoàn toàn)
1. **Phân tích yêu cầu**: Lấy Tên công ty (VD: Paritas), Tháng và Năm từ câu lệnh người dùng.
2. **Thực thi lệnh Python**: Bạn KHÔNG CẦN tìm file Tariff thủ công. Hãy sử dụng `uv run` để gọi script, script Python đã được trang bị thuật toán **Smart Scan**, nó sẽ tự lùng sục file Tariff có tên phù hợp nhất trong thư mục `tariff_dir`.
   ```bash
   uv run --with pandas --with openpyxl --with requests --with xlrd ~/.agent/skills/quyet-toan-efms/scripts/build_efms.py --company "Paritas" --month 4 --year 2026
   ```
   *(Nhớ thay đổi tham số `--company`, `--month`, `--year` cho phù hợp).*
3. **Trình bày kết quả**: 
   - Nếu script thành công: In ra đường dẫn file Excel báo cáo vừa tạo.
   - Nếu script báo lỗi không tìm thấy file Tariff: Báo lại cho người dùng biết để họ kiểm tra lại thư mục `tariff_dir` hoặc tên công ty đã đúng chưa.

## 3. Tính Năng Smart Scan & Xử Lý Trùng Lặp (Dành cho Agent hiểu)
Script Python sẽ:
- Lấy tên công ty bạn truyền vào (ví dụ `Paritas`).
- Quét toàn bộ file Excel (`.xls`, `.xlsx`) trong thư mục `tariff_dir`.
- Nếu tìm thấy **chính xác 1 file** chứa tên công ty, script sẽ tự động dùng file đó và lưu vào `company_map.json`.
- **ĐẶC BIỆT LƯU Ý:** Nếu script tìm thấy **nhiều file** cùng chứa tên công ty, nó sẽ in ra mã lỗi `[AMBIGUOUS_MATCH]` kèm danh sách các file. Khi gặp trường hợp này, Agent **BẮT BUỘC PHẢI DỪNG LẠI**, liệt kê danh sách file đó ra và hỏi người dùng xem file nào là chính xác. Sau khi người dùng chọn, Agent phải mở file `company_map.json` ra, lưu đường dẫn chính xác vào key tương ứng, rồi mới chạy lại lệnh `uv run`.

## 4. Các Logic Quy Ước (Hardcoded)
- Số tờ khai được lấy qua API và tách thành 3 dòng phí (`BO_COF_OPS`, `BO_NOF_OPS`, `BO_OTH_OPS`) và tự động **sắp xếp theo thứ tự**.
- **Cột A (HBL)**: Điền Số Bill (so_bill).
- **Partner Code**: Luôn là `VDTRUCKINGHD-000`.
- **Unit Price**: Tra cứu tự động cho TTHQ, mặc định 22000 cho truyền TK và 300000 cho Nâng Hạ.
- **Tỷ giá**: Luôn điền `24410`.
- **Invoice**: `00077` và ngày hiện tại.

## 5. Hỗ trợ dữ liệu Offline (Local Excel) & Các trường hợp đặc biệt
- **Đọc từ file Local**: Nếu khách hàng (như Haohua) gửi trực tiếp file Excel (Debit Note) thay vì API, hãy thêm tham số `--input-file "đường/dẫn/file.xlsx"`.
  - **Tuỳ chỉnh cột**: Hỗ trợ tuỳ chỉnh cột đọc bằng các tham số: `--bill-col <index>`, `--cont-col <index>`, `--tk-col <index>` (index bắt đầu từ 0). Ví dụ: Cột B = 1, Cột H = 7, Cột R = 17.
  - **Gộp tờ khai**: Script mặc định có chức năng **gộp nhiều tờ khai của cùng 1 bill** vào chung 1 shipment (ngăn cách bằng dấu phẩy) để tránh bị lặp phí container. Tính năng này được bật mặc định (`--group-tk`).
- **Chia nhỏ file (Split)**: Nếu file xuất ra quá dài, EFMS có thể lỗi khi import. Dùng thêm tham số `--split <số lượng>` (ví dụ: `--split 3`) để script tự động cắt đều dữ liệu thành nhiều file con (có đuôi `_01`, `_02`...).
- **Cảnh báo trùng lặp**: Kịch bản chạy sẽ tự động tô màu vàng các Số bill (HBL) bị trùng lặp trên file Excel xuất ra.
- **Đặc quyền Haohua Hàng Xuất**: Khi chạy cho công ty "Haohua Hàng Xuất" (hoặc từ khóa có chứa `haohua` và `xuất`):
  - Script sẽ **tự động loại bỏ hoàn toàn phí `BO_OTH_OPS` mặc định** và ép cứng phí **`BO_NOF_OPS` thành 190,000 VND / 1 cont**. Loại container mặc định là `40'DC`. Luồng mặc định là `Xanh`.
  - **Định dạng file Excel của Haohua Hàng Xuất** thường có cấu trúc: Số bill ở cột B (`--bill-col 1`), Số lượng cont ở cột H (`--cont-col 7`), Số lượng tờ khai ở cột R (`--tk-col 17`). Khi gọi script bắt buộc phải truyền các tham số này.
- **Đặc quyền Haohua Hàng Nhập**: Khi chạy cho công ty "Haohua Hàng Nhập" (hoặc từ khóa có chứa `haohua` và `nhập`), script sẽ **tự động loại bỏ hoàn toàn phí `BO_OTH_OPS` mặc định**. Phí **`BO_NOF_OPS` sẽ được set cứng thành 160,000 VND (luồng Xanh)** hoặc **190,000 VND (luồng Vàng)**. Đặc biệt, nếu là luồng Vàng, tự động thêm **Hiện trường mở tờ khai (BO_OTH_OPS) 300,000 VND / 1 lô (shipment)**.

## 6. Tính Năng Tải Hóa Đơn Điện Tử (Smartport API) - Tự động toàn tập cho HASITC & FORTUNE
Kịch bản tự động tải PDF hóa đơn từ Smartport dành cho Cảng Bình Dương (hoặc các cảng dùng Smartport):
- **Lệnh chạy**: `uv run --with playwright --with pandas --with openpyxl --with requests python scripts/download_invoices_api.py`
- **Cách thức hoạt động**:
  1. Script đọc Google Sheet, **tự động quét qua tất cả các sheet đang hiển thị (KHÔNG BỊ ẨN)** (VD: sheet của HASITC, FORTUNE,...).
  2. Ở mỗi sheet, tự động trích xuất "Tên Công Ty" từ tên sheet (Ví dụ: `FORTUNE - BẢNG KÊ - T5-2026` -> Tên công ty là `FORTUNE`).
  3. Lọc lấy các dòng có Nơi hạ cont là "Bình Dương", lấy `Loading Date` và tự tính khoảng thời gian (±30 ngày) để vượt qua giới hạn tra cứu 3 tháng của Smartport.
  4. Yêu cầu user tự thao tác **1 lần mồi** ở giao diện "Tra cứu e-Eir" trên trình duyệt.
  5. Bắt API mồi và tự động gọi API ngầm để tra cứu PinCode của toàn bộ danh sách container thu thập được.
  6. Sử dụng PinCode lấy được, trực tiếp gọi API download (`apismp/invoice/downloadInvPDF`) để kéo file PDF về máy trong nháy mắt.
- **Kết quả (Đầu ra chuẩn)**: Tất cả file PDF sẽ được tự động phân loại và lưu thẳng vào ổ đĩa theo cấu trúc: `d:\workspace-ai\HOA DON\<TÊN CÔNG TY>\THÁNG 5\` (Ví dụ: `HOA DON\FORTUNE\THÁNG 5\`). Mọi thứ hoàn toàn tự động, user không cần phải di chuyển file bằng tay!

## 7. Xử lý dữ liệu Buying (Đối soát Bảng Kê và Sản Lượng) - Dành riêng cho FORTUNE
Kịch bản tự động gộp và đối chiếu số tiền từ các file Sản Lượng nhiều tháng vào file Buying EFMS:
- **Lệnh chạy**: `uv run --with openpyxl python scripts/process_buying.py`
- **Các rule đối soát thông minh (Smart Matching)**:
  1. **Loại bỏ Hàng lẻ / Thiếu Cont**: Tự động quét Cột E (Số Cont) của Bảng Kê. Nếu ô bị trống hoặc chứa từ khóa `Hàng lẻ` / `Hang le`, kịch bản tự động chốt số tiền bằng `0` và đẩy sang danh sách cần kiểm tra tay (`_CHECK`).
  2. **Khớp số Bill bằng Số đuôi**: Tự động cắt bỏ các chữ cái ở đầu số Bill (VD: `EGLV235600578302` -> `235600578302`) để đối chiếu phần số đuôi. Điều này giúp khắc phục lỗi do con người khi nhập liệu thiếu chữ cái.
  3. **Dự phòng (Fallback) bằng Số Cont**: Nếu đối soát bằng số Bill thất bại, kịch bản sẽ tự động dùng Số Cont quét lại một lần nữa vào bảng Sản Lượng để tìm và cứu dữ liệu.
- **Kết quả (Tách file tự động)**:
  - File `_DONE.xlsx` lưu danh sách các bill đã đối soát thành công và gán đủ tiền (Cột Unit Price).
  - File `_CHECK.xlsx` lưu các bill không có data hoặc là hàng lẻ. Tại file này, hệ thống tự động điền **ngày thực hiện** (ngày hôm nay) vào cột M (Exchange Date) và cột P (Invoice Date) để người dùng tiện thao tác thủ công.

## 8. Xử lý Phiếu Cân Xe (Nông nghiệp E.H Việt Nam)
Khi người dùng yêu cầu đọc thông tin từ Phiếu Cân Xe (thường là file PDF dạng ảnh) và điền vào file Bảng Kê Excel (VD: `BẢNG KÊ EH VN HÀNG CONT THÁNG 5.xlsx`):
- **Bước 1 (Đọc PDF)**: Sử dụng công cụ `view_file` để mở PDF (hệ thống sẽ chụp màn hình cho Agent đọc). Dùng kỹ năng Vision trích xuất 2 thông tin chính: **Số xe** (VD: 72H05031) và **Trọng lượng hàng / TL Hàng** (VD: 32,740).
- **Bước 2 (Xử lý chuỗi)**: Format lại Số xe có dấu gạch ngang (VD: `72H05031` -> `72H-05031`).
- **Bước 3 (Ghi Excel bằng PowerShell)**: Dùng PowerShell COM Object (`Excel.Application`) để chèn dữ liệu vào dòng trống tiếp theo của sheet yêu cầu (thường là nối tiếp format của dòng 24).
  - **Cột A (Nội dung)**: `Cảng SPPSA về KCN Bàu Xéo - [Số xe]`
  - **Cột D (Số xe)**: `[Số xe]`
  - **Cột E (Trọng lượng)**: Ép kiểu `[double]` cho trọng lượng.
  - **Cột G (Đơn giá)**: Mặc định `180` (ép kiểu `[double]`).
  - Copy công thức tính từ dòng mẫu (dòng 24) cho các cột Thành tiền (H), VAT% (I), Tiền thuế (J).
  - **Cột K (Tổng tiền)**: Bắt buộc set công thức `=+H<row>+J<row>` (Ví dụ `=+H25+J25`) vì lệnh copy thường có thể bỏ sót cột K.
- **Lưu ý Code PowerShell**: Đảm bảo dùng `.Value2 = [double]...` để không bị lỗi Type Cast và giữ đúng định dạng số có dấu phẩy.

