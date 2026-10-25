# Chạy 30 file `test/` qua add-in trong Word ở máy ảo (25/10/2026)

Cách chạy: Word 16.0.20527 (Windows 11 ARM), add-in bản build `1053C5…` (đã gồm bản sửa NĐ 30 / Đảng / Đoàn).
Mỗi file: bản sao → bấm **Chuẩn hóa 1-Click** trên Ribbon (UI Automation), thiết lập tạm *Tự nhận diện*, cỡ 14 →
rà soát lại theo quy chuẩn mong đợi (theo tên file) → đo vị trí đường kẻ → xuất PDF. Thiết lập Ribbon đã trả về như cũ.

> **Cập nhật:** đã sửa lệch 1–4 (màu chữ về đen; bỏ đậm/nghiêng phủ cả đoạn ở nội dung, giữ nhấn mạnh từng chữ;
> trang ngang chỉ giữ khi có bảng nội dung, chuẩn hóa và rà soát dùng chung quy tắc; dấu phân cách dưới tên cơ quan Đảng `*`,
> Đoàn `***`). Chạy lại 30 file qua Ribbon (add-in `4B0069…`): 6 file ngang không bảng đã về A4 dọc, file 28 (có bảng) giữ
> ngang và không còn bị báo lỗi; 21 file còn 0 lỗi, còn lại chỉ thiếu cấp bậc hàm (06–14) và `GD5512-10` (27–30).
> **Cập nhật lần 2:** đã sửa lệch 5–10 (trích yếu công văn vào khối đầu dưới số ký hiệu, bỏ nhãn "công văn"; chữ trong bảng
> bỏ đậm/nghiêng/gạch chân phủ cả ô, độ rộng cột theo nội dung; "Chương I", tên chương viết thường được nhận và in hoa;
> "TÊN BÀI DẠY:", khoảng trắng dòng thông tin đầu, "Lưu:", "15/KH"; căn cứ ban hành chỉ trước Điều/đề mục đầu tiên).
> Chạy lại 30 file qua Ribbon (add-in `3CE50B…`): 21 file 0 lỗi, 9 file chỉ còn thiếu cấp bậc hàm người ký (lỗi của văn bản).
> Còn thấy: chữ cột đầu bảng KHBD giữ căn phải của file gốc.

## Đạt

- **Nhận diện quy chuẩn đúng 30/30.**
- Lề, khổ A4 (trừ trang ngang), phông Times New Roman, cỡ chữ các thành phần, số ký hiệu, ngày tháng, nơi nhận, khối chữ ký:
  đúng theo từng quy chuẩn; Đảng/Đoàn có `-----` / `-------`, chức vụ người ký không đậm; NĐ 30 khối đầu 12/13/13.
- **Đường kẻ**: mọi đường kẻ nằm ngay dưới dòng chữ (cách đáy chữ 2,2–2,4 pt; dưới tiêu đề Đảng/Đoàn 8,5 pt).
- Rà soát sau chuẩn hóa: 0 lỗi ở 17 file; lỗi còn lại xem bảng dưới.

## Còn lệch (theo mức độ, số file bị ảnh hưởng)

| # | Lệch | File | Ghi chú |
|---|---|---|---|
| 1 | **Màu chữ không về đen** (đỏ, xanh, tím… giữ nguyên) | 30/30 | HD 05 mục IV.1 ghi rõ màu đen; rà soát không báo |
| 2 | **Đoạn nội dung giữ đậm / nghiêng / gạch chân cả đoạn** | 30/30 | `Planner` không đặt đậm/nghiêng cho `Body`, `FormatChecker` bỏ qua `Body` → không sửa, không báo |
| 3 | **Trang ngang giữ nguyên nhưng rà soát báo lỗi khổ giấy** | 02, 07, 12, 16, 20, 24, 28 | `ApplyPageSetup` giữ hướng ngang (cho văn bản có bảng rộng), `CheckPage` đòi 210×297 → tự mâu thuẫn; văn bản ngắn không bảng vẫn bị để ngang |
| 4 | Dấu phân cách dưới tên cơ quan sai không sửa: Đảng ra `-----` thay `*`, Đoàn ra `*` thay `***` | 15, 17, 18 / 19, 21, 22 | `TextNormalizer` không sửa `DauSao`; rà soát không báo |
| 5 | Công văn Công an: dòng `V/v …` nằm giữa trang thay vì dưới số ký hiệu; dòng chữ lạc "công văn" | 11–14 | trích yếu công văn không được đưa vào khối đầu |
| 6 | Bảng: chữ trong ô giữ gạch chân + đậm + nghiêng; cột đầu quá hẹp, chữ xuống dòng từng từ | 27, 29 (cũng 28, 30) | `TableText` chỉ chỉnh cỡ chữ |
| 7 | NĐ 250: "Chương i" không đổi thành "Chương I"; "quy định chung" không in hoa, đậm | 23–26 | dòng thường sau "Chương" không nhận là tên chương |
| 8 | CV 5512: "Tổ : …", "Họ và tên giáo viên : …" — rà soát báo `GD5512-10` nhưng chuẩn hóa không sửa; "tên bài dạy" không in hoa | 27–30 | `ThongTinDau` không nằm trong phạm vi sửa chữ |
| 9 | "- Lưu : VT, VP." và "số 15 /KH-VP" giữ khoảng trắng thừa | 1–26 | `FixNoiNhanItem`, `FixSpacing` không xử lý trước ":" sau "Lưu" và trước "/" |
| 10 | Câu thân bài "Căn cứ kế hoạch số 15…" dưới Điều 1 bị in nghiêng như căn cứ ban hành | phần lớn | `Classifier`: "Căn cứ" trong 15 đoạn đầu nội dung đều là `CanCu` |
| 11 | Thiếu cấp bậc hàm người ký (đúng là lỗi của văn bản, rà soát báo, không tự điền) | 06–14 | đúng thiết kế |
