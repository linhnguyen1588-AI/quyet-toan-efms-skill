# Quy Trình Tự Động Hóa Tải Thông Báo Phí Cơ Sở Hạ Tầng (Thu Phí Hạ Tầng)

Tài liệu này ghi chú lại quy trình đã hoàn thiện để tra cứu và tải Thông báo phí cơ sở hạ tầng hàng hải (CSHT) từ hệ thống cảng, giúp tự động hóa quá trình lấy dữ liệu và đánh dấu file Excel.

## 1. Yêu cầu & Bối cảnh
- Hệ thống đích: Trang web Tra cứu biên lai thu phí CSHT.
- Dữ liệu đầu vào: File Excel (ví dụ `CSHT HAOHUA.xlsx`) chứa danh sách `Mã số tờ khai`.
- Yêu cầu:
  1. Tìm kiếm từng mã tờ khai.
  2. Bấm "Lấy thông báo" để tạo phí.
  3. Bấm tải file PDF (Thông báo phí/Biên lai).
  4. Lấy số Thông báo phí xuất ngược lại vào file Excel gốc (cột "Số thông báo phí (Hải quan)").
  5. Nếu tờ khai nào không có dữ liệu phí (số tiền = 000000000000) hoặc không tìm thấy, bôi nền **màu vàng** dòng đó trong Excel để dễ nhận biết.

## 2. Quy trình thao tác (Logic Script)
1. **Đăng nhập & Chờ người dùng cấu hình:**
   - Script mở trình duyệt Playwright, yêu cầu người dùng tự đăng nhập và điền CAPTCHA.
   - Script phát hiện khi người dùng vào trang "Tra cứu biên lai".
   - **(Quan trọng)** Script đếm ngược 15 giây để người dùng tự tay chọn khoảng thời gian (Từ ngày - Đến ngày) (VD: Từ 01/05/2026 đến hôm nay), do mặc định web có thể chỉ tìm theo ngày hiện tại dẫn đến không tìm thấy dữ liệu cũ.

2. **Xử lý từng tờ khai trong danh sách:**
   - **Xóa Pop-up (Dọn dẹp):** Bấm phím `Escape` và click vào các nút "Đóng" (`button:has-text("Đóng"):visible`) để dọn sạch mọi popup/modal ẩn từ lần chạy trước.
   - **Tìm kiếm:** Nhập `Mã tờ khai` vào ô tìm kiếm (`input[name="SO_TK"]`) và bấm click thẳng vào nút Tìm kiếm (`.btnSearch`). Wait một khoảng an toàn (4s) để bảng cập nhật.
   - **Kiểm tra dữ liệu:** Đọc HTML bảng kết quả:
     - Nếu không có dữ liệu -> Bôi vàng trong Excel.
     - Nếu có, đọc cột 10 (Số thông báo phí) và cột 11 (Số tiền).
     - Nếu số tiền là `000000000000` -> Tờ khai chưa có phí -> Bôi vàng trong Excel.
   - **Lấy thông báo (Tạo số TB):** Bấm nút "Lấy thông báo" (`.btnXemThongBaoNP`).
   - **Tải PDF:**
     - Đợi một lúc để modal xuất hiện.
     - Tìm nút "Tải thông báo nộp phí" hoặc "In thông báo" **ĐANG HIỂN THỊ** (`.btnPrintThongBaoNP, button:has-text("Tải thông báo nộp phí")` kết hợp filter `is_visible()`).
     - Bấm tải và lưu file PDF theo tên là `[Số Thông Báo].pdf`.
   - **Đóng Modal (Cực kỳ quan trọng):** Click nút `Đóng` của modal vừa mở để tránh việc nút Tải của các tờ khai sau bị che khuất hoặc nhầm lẫn DOM.

3. **Ghi lại Excel:**
   - Load file Excel, điền Số thông báo phí vào cột tương ứng.
   - Nếu có lỗi, dùng `openpyxl.styles.PatternFill(start_color="FFFFFF00")` để bôi vàng nguyên dòng.
   - Save file vào file `_Done.xlsx` (incremental save) sau mỗi lần xử lý để đề phòng lỗi đứt gánh.

## 3. Các bài học và Lỗi đã xử lý (Bug Fixes)
- **Lỗi tìm sai Nút Tải / Kẹt modal:**
  - Ban đầu script dùng `page.query_selector('.btnPrintThongBaoNP')` để tìm nút tải, nhưng vì web render nhiều modal đè lên nhau, script đã lấy nhầm nút của modal bị ẩn (ẩn nhưng vẫn tồn tại trong DOM). Giải pháp là dùng `query_selector_all` sau đó lọc ra nút `is_visible() == True`.
- **Lỗi không đóng Modal:**
  - Sau khi tải PDF xong, nếu không tắt modal, ô tìm kiếm và nút `.btnSearch` sẽ bị đè, dẫn đến lệnh Click tìm kiếm của tờ khai tiếp theo không có tác dụng. Kết quả là script dùng lại bảng dữ liệu của tờ khai cũ. Giải pháp: Luôn bấm `Đóng` sau khi tải và bấm `Escape` đầu mỗi vòng lặp.
- **Lỗi thiếu dải ngày tìm kiếm:**
  - Chỉ nhập số tờ khai là không đủ nếu ngày lập tờ khai nằm ngoài khoảng mặc định của hệ thống. Phải để người dùng thiết lập mốc ngày trước khi tự động chạy.
- **Lỗi không tìm thấy ô Checkbox tải:**
  - Từng cố gắng tìm `input[type="checkbox"]` để tải nhưng giao diện hiện tại không dùng checkbox mà click trực tiếp vào nút "Tải thông báo nộp phí" bên trong modal Lấy thông báo.

## 4. Kiến trúc "Nồi đồng cối đá" (V3) và Khả năng phục hồi
Trong quá trình vận hành với số lượng lớn (hàng trăm tờ khai), hệ thống Hải Quan/Cảng thường xuyên bị quá tải, dẫn đến timeout hoặc văng phiên đăng nhập. Script (V3) đã được nâng cấp mạnh mẽ với các tính năng sau:

### 4.1 Cơ chế Resume (Nhớ bài) thông minh
- **Vấn đề:** Không thể bắt người dùng chạy lại từ đầu nếu rớt mạng ở tờ khai thứ 200.
- **Giải pháp:** Đầu mỗi phiên chạy, tool đọc file `_Done.xlsx` để quét xem tờ khai nào đã có `Số TB phí` hoặc đã được `bôi vàng` (xác nhận không có phí).
- **Cải tiến Cross-check (Kiểm tra chéo):** Tool không chỉ tin vào Excel, mà còn kiểm tra chéo xem `file PDF` tương ứng trong thư mục `PDFs` có thực sự tồn tại hay không. Nếu có số trong Excel nhưng rớt mất file PDF (do lỗi Permission denied hoặc rớt mạng khi tải), tool sẽ coi như tờ khai đó chưa hoàn tất và tự động chạy lại để tải bù.

### 4.2 Xử lý Timeout và Chống Crash
- **Vấn đề:** Web mất quá lâu (hơn 30 giây) để trả về bảng dữ liệu của 1 tờ khai.
- **Giải pháp:** Thay vì đánh dấu là lỗi hoặc làm crash script, tool bắt lỗi `TimeoutError`, in cảnh báo `Bỏ qua tờ khai này để thử lại sau` và lẳng lặng đi tiếp. Lần chạy lệnh sau, nhờ cơ chế Resume, tool sẽ bốc đúng tờ bị timeout này ra thử lại.

### 4.3 Tự động bắt lỗi văng Session (Hết hạn phiên)
- **Vấn đề:** Đang chạy ngon thì web tự động đá văng ra màn hình đăng nhập.
- **Giải pháp:** Ở đầu mỗi vòng lặp, tool kiểm tra URL (`page.url`). Nếu phát hiện chuỗi `dang-nhap`, tool sẽ:
  1. Tạm dừng tiến trình (`Pause`).
  2. In cảnh báo lớn ra màn hình bắt người dùng đăng nhập lại.
  3. Sau khi người dùng đăng nhập xong, tool sẽ tự động nhặt lại đúng tờ khai đang bị bỏ dở và chạy tiếp, không rớt nhịp nào.

### 4.4 Ghi dữ liệu an toàn (Incremental Save)
- Để tránh việc ghi vào Excel từng dòng gây nghẽn I/O (chậm ổ cứng), tool lưu tạm dữ liệu vào RAM (`tb_map`).
- **Ghi định kỳ:** Cứ mỗi 10 tờ khai xử lý xong, tool sẽ xả dữ liệu xuống file `_Done.xlsx` một lần.
- **Bảo vệ rớt dữ liệu:** Sau khi vòng lặp kết thúc (hoặc khi chạy qua những tờ khai bị bỏ qua cuối cùng), tool luôn có bước `Final save` để đảm bảo không một kết quả nào bị rớt lại dù xảy ra lỗi `Permission denied` ở các vòng trước đó.

## Tổng kết
Quy trình này hiện tại đã đạt độ ổn định rất cao, hoàn toàn có thể cắm máy chạy qua đêm hoặc xử lý những danh sách hàng nghìn tờ khai với độ trễ và tỷ lệ lỗi mạng của hệ thống gốc được triệt tiêu hoàn toàn nhờ khả năng **vét cạn** của tính năng Resume thông minh.
