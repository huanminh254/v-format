# Bộ 30 DOCX sai thể thức nặng

Bộ dữ liệu này dùng để kiểm thử khả năng tự nhận diện quy chuẩn và phát hiện lỗi của V-FORMAT. Mỗi file vẫn giữ từ khóa neo để xác định đúng một trong 7 quy chuẩn, nhưng cố ý sai đồng thời ở nhiều thành phần: khổ giấy, lề, phông, cỡ chữ, kiểu chữ, căn lề/thụt dòng, tiêu đề, số ký hiệu/ngày tháng, chữ ký/nơi nhận và số trang.

## Phân bố

| Quy chuẩn | Số file | Neo nhận diện |
|---|---:|---|
| Nghị định 30/2020/NĐ-CP | 5 | Nhánh mặc định, không có neo của 6 hệ chuyên biệt |
| Thông tư 68/2023/TT-BQP | 5 | `BỘ QUỐC PHÒNG` |
| Văn bản hợp nhất 08/VBHN-BCA | 4 | `CÔNG AN TỈNH MÔ PHỎNG` |
| Hướng dẫn 05-HD/VPTW | 4 | `ĐẢNG CỘNG SẢN VIỆT NAM` |
| Hướng dẫn 29-HD/TWĐTN-VP | 4 | `ĐOÀN TNCS HỒ CHÍ MINH` |
| Pháp điển và hợp nhất văn bản QPPL | 4 | `VĂN BẢN HỢP NHẤT` |
| Công văn 5512/BGDĐT-GDTrH | 4 | `KẾ HOẠCH BÀI DẠY` |

## Kết quả kiểm tra nền

- 30/30 file là gói DOCX/XML hợp lệ và mở lại được bằng `python-docx`.
- CLI V-FORMAT hiện tại tự nhận diện đúng 30/30 quy chuẩn.
- Mỗi file tạo ra từ 81 đến 96 lỗi trước chuẩn hóa.
- Mỗi file phủ từ 7 đến 10 nhóm mã lỗi.
- Đã render và kiểm tra trực quan toàn bộ 54 trang.

Các nhóm mã lỗi: `01` khổ giấy, `02` lề trang, `03` phông chữ, `04` cỡ chữ, `05` kiểu chữ, `06` căn lề/thụt dòng, `07` quốc hiệu/tiêu ngữ/tiêu đề, `08` số ký hiệu/ngày tháng, `09` chữ ký/nơi nhận, `10` thành phần/khoảng trắng/số trang.

## Danh sách file và mức lỗi hiện tại

| # | File | Quy chuẩn mong đợi | Lỗi | Nhóm mã phát hiện |
|---:|---|---|---:|---|
| 01 | `01-nd30-letter-le-hep-font-arial.docx` | NĐ30 | 93 | 01–10 |
| 02 | `02-nd30-a4-ngang-font-calibri.docx` | NĐ30 | 93 | 01–10 |
| 03 | `03-nd30-a5-font-courier-co-lon.docx` | NĐ30 | 92 | 01–10 |
| 04 | `04-nd30-legal-font-tahoma-co-nho.docx` | NĐ30 | 94 | 01–10 |
| 05 | `05-nd30-trang-vuong-font-verdana.docx` | NĐ30 | 92 | 01–10 |
| 06 | `06-tt68-letter-le-hep-font-arial.docx` | TT68 | 92 | 01–10 |
| 07 | `07-tt68-a4-ngang-font-calibri.docx` | TT68 | 91 | 01–10 |
| 08 | `08-tt68-a5-font-courier-co-lon.docx` | TT68 | 93 | 01–10 |
| 09 | `09-tt68-legal-font-tahoma-co-nho.docx` | TT68 | 92 | 01–10 |
| 10 | `10-tt68-trang-vuong-font-verdana.docx` | TT68 | 94 | 01–10 |
| 11 | `11-ca08-letter-le-hep-font-arial.docx` | CA08 | 90 | 01–10 |
| 12 | `12-ca08-a4-ngang-font-calibri.docx` | CA08 | 89 | 01–10 |
| 13 | `13-ca08-a5-font-courier-co-lon.docx` | CA08 | 92 | 01–10 |
| 14 | `14-ca08-legal-font-tahoma-co-nho.docx` | CA08 | 91 | 01–10 |
| 15 | `15-hd05-letter-le-hep-font-arial.docx` | HD05 | 90 | 01–06, 08–10 |
| 16 | `16-hd05-a4-ngang-font-calibri.docx` | HD05 | 89 | 01–06, 08–10 |
| 17 | `17-hd05-a5-font-courier-co-lon.docx` | HD05 | 91 | 01–06, 08–10 |
| 18 | `18-hd05-legal-font-tahoma-co-nho.docx` | HD05 | 89 | 01–06, 08–10 |
| 19 | `19-hd29-letter-le-hep-font-arial.docx` | HD29 | 89 | 01–06, 08–10 |
| 20 | `20-hd29-a4-ngang-font-calibri.docx` | HD29 | 89 | 01–06, 08–10 |
| 21 | `21-hd29-a5-font-courier-co-lon.docx` | HD29 | 89 | 01–06, 08–10 |
| 22 | `22-hd29-legal-font-tahoma-co-nho.docx` | HD29 | 88 | 01–06, 08–10 |
| 23 | `23-nd250-letter-le-hep-font-arial.docx` | NĐ250 | 95 | 01–10 |
| 24 | `24-nd250-a4-ngang-font-calibri.docx` | NĐ250 | 91 | 01–10 |
| 25 | `25-nd250-a5-font-courier-co-lon.docx` | NĐ250 | 96 | 01–10 |
| 26 | `26-nd250-legal-font-tahoma-co-nho.docx` | NĐ250 | 95 | 01–10 |
| 27 | `27-cv5512-letter-le-hep-font-arial.docx` | CV5512 | 83 | 01–06, 10 |
| 28 | `28-cv5512-a4-ngang-font-calibri.docx` | CV5512 | 81 | 01–06, 10 |
| 29 | `29-cv5512-a5-font-courier-co-lon.docx` | CV5512 | 82 | 01–06, 10 |
| 30 | `30-cv5512-legal-font-tahoma-co-nho.docx` | CV5512 | 82 | 01–06, 10 |

Số lỗi trong bảng là đường cơ sở theo mã nguồn hiện có tại thời điểm tạo bộ test. Khi thuật toán nhận diện hoặc quy tắc kiểm tra thay đổi, số lỗi có thể thay đổi; quy chuẩn mong đợi và neo nhận diện vẫn giữ nguyên để phục vụ hồi quy.
