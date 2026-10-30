
Module 1: Thiết lập "Căn cứ số" (Terminal)


1. Lệnh khởi tạo cây thư mục 3 cấp bằng một dòng lệnh duy nhất (Một lệnh Single Compact)

Sử dụng câu lệnh PowerShell tối ưu hóa với kỹ thuật mảng lồng nhau để tạo toàn bộ cấu trúc cây thư mục:
powershell
New-Item -ItemType Directory -Path ( "docs", "scripts", "data", "backup" | ForEach-Object { "greentech-system/$_" } )
2. Thao tác với đường dẫn tương đối (Relative Path) và tạo file cấu hình ẩn .env

Di chuyển vào thư mục gốc bằng đường dẫn tương đối và tạo tệp cấu hình ẩn bảo mật chứa thông tin giả định:
powershell
cd greentech-system
New-Item -ItemType File -Name ".env" -Value "DB_PASSWORD=GreenTechSecure2026`nAPI_KEY=IoT_Sensor_Farm_Secret"
# Thiết lập thuộc tính ẩn (Hidden) cho file .env trên môi trường Windows
(Get-Item ".env").Attributes = 'Hidden'
Module 2: Robot lọc Log tự động (Advanced Scripting)

Tạo tệp văn bản đặt tên là auto_report.bat nằm trong thư mục scripts/. Thao tác xử lý dữ liệu lớn qua bộ lọc đường ống Pipe (|) giúp hệ thống truyền thẳng luồng ký tự sang bộ nhớ đệm, không nạp toàn bộ file 500MB vào RAM để tránh tràn bộ nhớ.
Nội dung tệp auto_report.bat:
batch
@echo off
:: BƯỚC AUDIT (THẨM ĐỊNH): Sử dụng biến cục bộ và kiểm tra file trước khi xử lý, ngăn chặn lệnh xóa thô bạo del /f /s /q
setlocal enabledelayedexpansion

set "LOG_FILE=..\data\log-he-thong.log"
set "REPORT_FILE=..\backup\bao-cao-su-co.txt"

echo ===================================================
echo [AUDIT] DANG KIEM TRA HE THONG TRUOC KHI LOC LOG...
echo ===================================================

if not exist "%LOG_FILE%" (
    echo [ERROR] Khong tim thay file log-he-thong.log tai thu muc data/
    exit /b 1
)

:: Tiến hành tạo tiêu đề báo cáo chuyên nghiệp
echo =================================================== > "%REPORT_FILE%"
echo            BAO CAO SU CO TRONG HE THONG            >> "%REPORT_FILE%"
echo            Ngay thuc hien: %date% %time%           >> "%REPORT_FILE%"
echo =================================================== >> "%REPORT_FILE%"
echo. >> "%REPORT_FILE%"

:: ROBOT LOC LOG: Sử dụng toán tử Pipe (|) phối hợp findstr lọc lỗi, chèn nối bằng toán tử >>
echo [PROCESSING] Robot dang loc cac loi CRITICAL va ERROR...
type "%LOG_FILE%" | findstr /I "CRITICAL ERROR" >> "%REPORT_FILE%"

echo =================================================== >> "%REPORT_FILE%"
echo [SUCCESS] Bao cao da duoc xuat vao thu muc backup/bao-cao-su-co.txt
echo ===================================================
pause
Module 3: Tài liệu hướng dẫn vận hành (Markdown & Word)


1. Cấu trúc nội dung tệp README.md (Lưu tại docs/README.md)

markdown
# Hệ thống quản lý hạ tầng GreenTech IoT System

## Hướng dẫn chạy script lọc Log tự động
Để khởi chạy Robot quét lỗi và tự động kết xuất báo cáo sự cố hệ thống, lập trình viên thực hiện các bước sau:
1. Mở cửa sổ dòng lệnh `Command Prompt` (cmd) hoặc `PowerShell`.
2. Di chuyển đường dẫn vào thư mục chứa script:
   ```bash
   cd greentech-system/scripts
   ```
3. Chạy tệp tin thực thi:
   ```bash
   .\auto_report.bat
   ```

## Checklist kiểm tra thiết bị trạm IoT hàng ngày
- [ ] Kiểm tra trạng thái cấp nguồn của 100 trạm cảm biến.
- [ ] Xác thực kết nối mạng diện rộng Mbps ổn định.
- [ ] Quét lỗi dung lượng bộ nhớ SSD/HDD tại Sheet `trang-thai-tram`.
- [ ] Khởi chạy script audit tệp nhật ký định kỳ.

## Bảng thông số phần cứng cơ bản

| Loại thiết bị | Dung lượng RAM tối thiểu | Đơn giá tham khảo | Loại Storage hỗ trợ |
| :--- | :--- | :--- | :--- |
| Gateway Center | 16 GB | 1.500.000 VND | SSD |
| Edge Sensor Node | 4 GB | 450.000 VND | HDD |
2. Quy chuẩn xây dựng Đặc tả hệ thống srs-he-thong.docx

• Cấu trúc Styles: Áp dụng trực tiếp thanh công cụ Word Styles. Gán nhãn thuộc tính Heading 1 cho các mục lớn (1. Mục tiêu, 2. Mô tả bài toán), Heading 2 cho phân mục con (A. Kiến trúc lưu trữ, B. Database Thiết kế) và Heading 3 cho các module chức năng.
• Auto TOC: Di chuyển con trỏ lên trang đầu tiên \(\rightarrow \) Chọn Tab References \(\rightarrow \) Chọn Table of Contents \(\rightarrow \) Chọn mẫu Automatic Table 1 để tạo mục lục tự động đồng bộ theo cấu trúc Styles.

Module 4: Đối soát tài nguyên (Excel Database Mindset)

Để giải quyết triệt để Bẫy lỗi khoảng trắng thừa khi sao chép mã gây lỗi hiển thị #N/A, bạn thiết lập cấu trúc gọi hàm lồng nhau tại Sheet trang-thai-tram.
• Cấu trúc dữ liệu yêu cầu: Vùng bảng đối soát tuyệt đối không được dùng tính năng Merge Cell hay để dòng trống ngắt quãng, giữ bảng thô thuần túy để tính năng lọc dữ liệu AutoFilter phân tích chính xác.
• Công thức tại ô đối soát tên thiết bị (Sheet trang-thai-tram): Giả sử mã linh kiện cần tra nằm tại ô B2 thuộc Sheet trang-thai-tram, và bảng danh mục linh kiện nguồn nằm tại Sheet danh-muc-linh-kien từ vùng A2:D101.
Bạn nhập công thức chuẩn xác sau vào ô tính ten_thiet_bi:
excel
=VLOOKUP(TRIM(B2), 'danh-muc-linh-kien'!$A$2:$D$101, 2, FALSE)
• Giải nghĩa cơ chế hoạt động: Hàm TRIM(B2) đóng vai trò như một bộ lọc làm sạch dữ liệu đầu vào, tự động cắt bỏ mọi ký tự khoảng trắng thừa ở hai đầu mã linh kiện. Việc cố định vùng tham chiếu bằng ký tự $ ($A$2:$D$101) đảm bảo khi bạn kéo AutoFill công thức xuống 100 dòng, tọa độ danh mục linh kiện không bị chạy lệch pha dữ liệu.

Phân tích Khoa học thông tin & Bẫy Hardware


1. Mô hình xử lý thông tin IPO & Tư duy hệ thống DIKW

• Mô hình IPO (Input - Process - Output):
	• Input: Tệp tin log hệ thống thô dung lượng 500MB và bảng mã thông số linh kiện phần cứng.
	• Process: Sử dụng lệnh lọc chuỗi văn bản thông qua đường ống dẫn Pipe (findstr) kết hợp thuật toán đối soát tìm kiếm chuỗi sạch (VLOOKUP + TRIM).
	• Output: Tệp báo cáo sự cố được phân loại bao-cao-su-co.txt và cơ sở dữ liệu Asset Database đồng bộ hoàn chỉnh.
• Tư duy DIKW (Data - Information - Knowledge - Wisdom):
	• Data (Dữ liệu thô): Các dòng ký tự log rời rạc và các mã linh kiện chưa được xử lý.
	• Information (Thông tin): Lọc ra chính xác các dòng chữ ghi nhận lỗi ERROR 404, TIMEOUT hoặc mã thiết bị tương thích.
	• Knowledge (Tri thức): Nhận diện được trạm cảm biến nào đang có tần suất phát sinh lỗi cao nhất ngoài trang trại dựa trên bảng đối soát liên kết.
	• Wisdom (Trí tuệ/Giải pháp): Kỹ sư DevOps đưa ra quyết định hạ tầng: Bảo trì sớm cụm phần cứng bị lỗi, nâng cấp cấu hình RAM/SSD để hệ thống GreenTech vận hành không gián đoạn.

2. Tính toán Hardware quy đổi tốc độ mạng (Mbps sang MB/s)

• Quy tắc cốt lõi: Hệ số chuyển đổi giữa đơn vị Megabit (Mb) và Megabyte (MB) là 8 (1 Byte = 8 Bits).
• Bài toán: Giả sử tốc độ băng thông mạng truyền tải thực tế tại trang trại đạt 80 Mbps.
• Công thức quy đổi:
\(\text{Tốc độ tải xuống thực tế (MB/s)}=\frac{80\text{\ Mbps}}{8}=10\text{\ MB/s}\)
• Thời gian tải file log 500MB tối thiểu:
\(\text{Thời gian}=\frac{500\text{\ MB}}{10\text{\ MB/s}}=50\text{\ giây}\)
