# Bộ 30 DOCX lỗi nặng dựa trên “BỘ CÔNG AN.docx”

Bộ dữ liệu dùng để kiểm thử nhận diện quy chuẩn và chuẩn hóa thể thức của V-FORMAT. Nội dung quyết định, căn cứ và danh sách năm cán bộ được lấy từ tệp nguồn; phần neo nhận diện được điều chỉnh có kiểm soát để phủ đủ 7 hệ quy chuẩn.

## Nguồn và phạm vi

- Nguồn: `/Users/huan/Downloads/BỘ CÔNG AN.docx`
- SHA-256 nguồn: `1286bfe79bc18f174b478895c50a49bff46b7e8bb2bfe5a247b8602585cdd6c5`
- 30/30 tệp nhận diện đúng quy chuẩn mong đợi.
- 30/30 tệp là gói DOCX hợp lệ, mở lại được và đạt schema OOXML Office 2013.
- Mỗi tệp có từ 151 đến 223 lỗi theo CLI hiện tại.
- Tổng cộng 140 trang đã render; không có trang trắng hoặc lỗi kết xuất.

## Lỗi cố ý

- Khổ Letter, A4 ngang, A5, Legal và trang vuông 200 mm.
- Lề quá hẹp, quá rộng hoặc lệch nhau; số trang đặt sai đầu/chân trang.
- Arial, Calibri, Courier New, Tahoma, Verdana; cỡ chữ cực nhỏ hoặc quá lớn.
- Chữ màu, đậm/nghiêng/gạch chân/gạch xoá/in hoa nhỏ lẫn lộn.
- Căn trái/phải/giữa/phân tán, thụt đầu dòng âm hoặc quá sâu, khoảng cách đoạn sai.
- Quốc hiệu, tiêu ngữ, số ký hiệu, ngày tháng, chức vụ, nơi nhận và cấu trúc bài dạy bị làm sai.
- Bảng nhân sự cố ý quá rộng, màu sắc mạnh, chiều cao hàng và lề ô không đều.

## Danh sách tệp

| # | File | Chuẩn | Lỗi | Nhóm mã | Trang |
|---:|---|---|---:|---|---:|
| 01 | `01-nd30-bca-letter-le-sieu-hep-arial.docx` | NĐ30 | 196 | 01, 02, 03, 04, 05, 06, 07, 08, 09, 10 | 5 |
| 02 | `02-nd30-bca-a4-ngang-le-lech-calibri.docx` | NĐ30 | 164 | 02, 03, 04, 05, 06, 09, 10 | 2 |
| 03 | `03-nd30-bca-a5-le-am-courier-co-lon.docx` | NĐ30 | 189 | 01, 02, 03, 04, 05, 06, 07, 08, 09, 10 | 8 |
| 04 | `04-nd30-bca-legal-le-rong-tahoma-co-nho.docx` | NĐ30 | 164 | 01, 02, 03, 04, 05, 06, 09, 10 | 2 |
| 05 | `05-nd30-bca-vuong-200mm-verdana-mau-sac.docx` | NĐ30 | 199 | 01, 02, 03, 04, 05, 06, 07, 08, 09, 10 | 8 |
| 06 | `06-tt68-bca-letter-le-sieu-hep-arial.docx` | TT68 | 203 | 01, 02, 03, 04, 05, 06, 07, 08, 09, 10 | 5 |
| 07 | `07-tt68-bca-a4-ngang-le-lech-calibri.docx` | TT68 | 156 | 02, 03, 04, 05, 06, 09, 10 | 2 |
| 08 | `08-tt68-bca-a5-le-am-courier-co-lon.docx` | TT68 | 193 | 01, 02, 03, 04, 05, 06, 07, 08, 09, 10 | 7 |
| 09 | `09-tt68-bca-legal-le-rong-tahoma-co-nho.docx` | TT68 | 156 | 01, 02, 03, 04, 05, 06, 09, 10 | 2 |
| 10 | `10-tt68-bca-vuong-200mm-verdana-mau-sac.docx` | TT68 | 197 | 01, 02, 03, 04, 05, 06, 07, 08, 09, 10 | 7 |
| 11 | `11-ca08-bca-letter-le-sieu-hep-arial.docx` | CA08 | 201 | 01, 02, 03, 04, 05, 06, 07, 08, 09, 10 | 5 |
| 12 | `12-ca08-bca-a4-ngang-le-lech-calibri.docx` | CA08 | 160 | 02, 03, 04, 05, 06, 09, 10 | 3 |
| 13 | `13-ca08-bca-a5-le-am-courier-co-lon.docx` | CA08 | 191 | 01, 02, 03, 04, 05, 06, 07, 08, 09, 10 | 8 |
| 14 | `14-ca08-bca-legal-le-rong-tahoma-co-nho.docx` | CA08 | 162 | 01, 02, 03, 04, 05, 06, 09, 10 | 2 |
| 15 | `15-hd05-bca-letter-le-sieu-hep-arial.docx` | HD05 | 198 | 01, 02, 03, 04, 05, 06, 07, 08, 09, 10 | 5 |
| 16 | `16-hd05-bca-a4-ngang-le-lech-calibri.docx` | HD05 | 157 | 02, 03, 04, 05, 06, 09, 10 | 2 |
| 17 | `17-hd05-bca-a5-le-am-courier-co-lon.docx` | HD05 | 185 | 01, 02, 03, 04, 05, 06, 07, 08, 09, 10 | 9 |
| 18 | `18-hd05-bca-legal-le-rong-tahoma-co-nho.docx` | HD05 | 161 | 01, 02, 03, 04, 05, 06, 09, 10 | 2 |
| 19 | `19-hd29-bca-letter-le-sieu-hep-arial.docx` | HD29 | 189 | 01, 02, 03, 04, 05, 06, 07, 08, 09, 10 | 5 |
| 20 | `20-hd29-bca-a4-ngang-le-lech-calibri.docx` | HD29 | 154 | 02, 03, 04, 05, 06, 09, 10 | 2 |
| 21 | `21-hd29-bca-a5-le-am-courier-co-lon.docx` | HD29 | 196 | 01, 02, 03, 04, 05, 06, 07, 08, 09, 10 | 10 |
| 22 | `22-hd29-bca-legal-le-rong-tahoma-co-nho.docx` | HD29 | 151 | 01, 02, 03, 04, 05, 06, 09, 10 | 2 |
| 23 | `23-nd250-bca-letter-le-sieu-hep-arial.docx` | NĐ250 | 201 | 01, 02, 03, 04, 05, 06, 07, 08, 09, 10 | 5 |
| 24 | `24-nd250-bca-a4-ngang-le-lech-calibri.docx` | NĐ250 | 156 | 02, 03, 04, 05, 06, 07, 09, 10 | 3 |
| 25 | `25-nd250-bca-a5-le-am-courier-co-lon.docx` | NĐ250 | 189 | 01, 02, 03, 04, 05, 06, 07, 08, 09, 10 | 9 |
| 26 | `26-nd250-bca-legal-le-rong-tahoma-co-nho.docx` | NĐ250 | 163 | 01, 02, 03, 04, 05, 06, 07, 09, 10 | 2 |
| 27 | `27-cv5512-bca-letter-le-sieu-hep-arial.docx` | CV5512 | 223 | 01, 02, 03, 04, 05, 06, 09, 10, 11 | 5 |
| 28 | `28-cv5512-bca-a4-ngang-le-lech-calibri.docx` | CV5512 | 177 | 02, 03, 04, 05, 06, 09, 10, 11 | 3 |
| 29 | `29-cv5512-bca-a5-le-am-courier-co-lon.docx` | CV5512 | 220 | 01, 02, 03, 04, 05, 06, 09, 10, 11 | 7 |
| 30 | `30-cv5512-bca-legal-le-rong-tahoma-co-nho.docx` | CV5512 | 181 | 01, 02, 03, 04, 05, 06, 09, 10, 11 | 3 |

Số lỗi là đường cơ sở theo mã nguồn tại thời điểm tạo. Khi thuật toán nhận diện hoặc chuẩn hóa thay đổi, số lượng có thể thay đổi; neo quy chuẩn và tệp nguồn vẫn cố định để phục vụ hồi quy.
