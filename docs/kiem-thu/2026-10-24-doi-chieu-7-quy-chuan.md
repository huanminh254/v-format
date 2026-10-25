# Đối chiếu V-FORMAT với 7 hệ quy chuẩn gốc (24/10/2026)

> **Cập nhật 25/10/2026 — sửa các mục còn lại** (unit test trong `tests/VFormat.Core.Tests/QuyChuanTests.cs`; **chưa chạy lại
> trên Word ở máy ảo**):
> - **A1** khối đầu gõ tab: nhận Quốc hiệu/Tiêu ngữ/ngày tháng ở phần phải, tách phần trái (cơ quan, số ký hiệu), nối tên cơ quan
>   bị xuống dòng, dàn thành bảng 2 cột; rà soát không còn báo sai "Thiếu Quốc hiệu", báo `-07` "gõ hai cột bằng tab".
> - **A2** chữ ký đặt bằng thụt lề (≥ 5 cm): nhận chức vụ, "(Ký, ghi rõ họ tên)", họ tên; giữ thụt lề, căn giữa; đề mục
>   "V. Ý KIẾN…" không còn bị nhận là chức vụ. **A3** đề mục đánh số tự động: đọc `ListString` (Word) / `numbering.xml` (.docx).
>   **A4** đã sửa từ trước. **A5** ghi chú biểu mẫu → loại `GhiChuMau`, giữ nguyên. **A6** đoạn danh sách không bị kiểm thụt dòng.
>   Thêm: câu văn ngay sau trích yếu (kết thúc bằng dấu chấm) là nội dung, không còn là trích yếu dòng 2.
> - **B.3** Công an: cấp bậc hạ sĩ quan, chiến sĩ; chức danh khoa học/danh hiệu trước cấp bậc ("PGS.TS Thiếu tướng …").
> - **B.6** NĐ 250/2026 (đọc lại Phụ lục tr. 3–8, 58): đổi tên thành "Nghị định 250/2026/NĐ-CP"; "VĂN BẢN HỢP NHẤT" chuyển vào
>   ô trái trên số ký hiệu, không kẻ dưới; rà soát số có năm + `VBHN`; phần ký xác thực (cơ quan ký xác thực có kẻ dưới,
>   "XÁC THỰC VĂN BẢN HỢP NHẤT", ngày tháng, chức vụ, "(Chữ ký và dấu)", họ tên) dàn vào cột phải; đường kẻ ngang nét liền
>   (viền dưới đoạn nội dung cuối) ngăn phần ký xác thực. Mẫu thư viện viết lại theo ví dụ 2 và mục XIII.
>   Thông số trang/phông/cỡ vẫn theo NĐ 30 vì văn bản QPPL làm nền chưa có trong bộ nguồn.
> - **B.7** CV 5512: kiểm khung Phụ lục IV (mã `GD5512-11`); lỗi lề/phông/cỡ chữ hạ xuống mức khuyến nghị (Info);
>   mẫu thư viện dùng đúng chữ của Phụ lục IV ("I. Mục tiêu", tên đầy đủ Hoạt động 2).
> - **B.2** TT 68: tra lại 25/10/2026 vẫn không có toàn văn công khai kiểm chứng được — giữ nguyên, chưa xác nhận.

> **Cập nhật cùng ngày:** đã sửa toàn bộ sai lệch thông số ở mục B.1 (NĐ 30), B.4 (Đảng) và B.5 (Đoàn), có unit test
> cho từng điểm; chạy lại trên Word ở máy ảo (mẫu Quyết định, Nghị quyết Đảng, Báo cáo Đoàn: còn 0 lỗi; định dạng thực tế khớp
> bảng dưới). Mục A (engine), B.3 (chức danh trước cấp bậc), B.6 (NĐ 250), B.7 (CV 5512) **chưa sửa**.
> Đường kẻ: khi bấm **Chuẩn hóa 1-Click trên Ribbon** (add-in chạy trong Word, tự nhận diện) đường kẻ dưới tên cơ quan,
> tiêu ngữ, tiêu đề Đảng/Đoàn và trích yếu nằm đúng ngay dưới dòng chữ; Đảng/Đoàn chèn đúng `-----` / `-------`.
> Đường kẻ chỉ lệch khi gọi `Formatter.Run` từ một tiến trình khác qua COM (cách kiểm thử ở mục dưới) — không ảnh hưởng người dùng,
> nhưng vị trí đường kẻ trong các tệp kiểm thử tạo theo cách đó không đáng tin.

## Cách kiểm thử

- Máy ảo UTM Windows 11 ARM, Word 16.0.20527 (Microsoft 365), add-in V-FORMAT 1.1.0 bản build 20:35
  (SHA-256 `E5215440…D032`, gồm cả 4 bản sửa TT 68 / Đảng).
- Tệp thử: `báo cáo T9.docx` (Mẫu BL2, văn bản Công an thật, 41 đoạn, 2 trang). Với mỗi quy chuẩn: tạo tài liệu mới
  trong Word đang chạy → chèn tệp gốc → gọi thẳng `Formatter.Run` (Chuẩn hóa 1-Click) và `Review.Run` (Rà soát) của
  DLL đã cài, ép quy chuẩn tương ứng → đọc lại định dạng thật từng đoạn bằng Word → lưu `.docx` + `.pdf`.
  Tự nhận diện chọn **Công an (VBHN 08)** — đúng.
- Văn bản gốc đối chiếu: `docs/quy-chuan/` (NĐ 30/2020 Phụ lục I; VBHN 08/2021 Điều 9, 13; HD 05-HD/VPTW **ngày 27/5/2026**;
  HD 29-HD/TWĐTN-VP 2013 Mẫu 3; NĐ **250/2026**/NĐ-CP Phụ lục; CV 5512 Phụ lục IV). TT 68/2023/TT-BQP chưa có toàn văn.
- Số trang ghi trong ngoặc là trang của tệp PDF nguồn.

Kết quả rà soát của chính V-FORMAT sau chuẩn hóa: còn 7 lỗi ở 6 quy chuẩn, 4 lỗi ở CV 5512. Con số này **không phản ánh**
các lỗi bố cục bên dưới — rà soát không thấy chúng.

## A. Lỗi chung của engine (xảy ra ở cả 7 quy chuẩn trên báo cáo T9)

| # | Hiện tượng sau chuẩn hóa | Nguyên nhân trong code |
|---|---|---|
| A1 | Khối đầu bị vỡ thêm: dòng "CỤC CẢNH SÁT GIAO ⇥ CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM / THÔNG ⇥ Độc lập - Tự do - Hạnh phúc" bị căn đều, thụt 1 cm; không dàn bảng 2 cột, không kẻ dưới tiêu ngữ, cơ quan. Rà soát báo **sai** "Thiếu Quốc hiệu". | Cơ quan và Quốc hiệu nằm chung một đoạn, cách nhau bằng tab; `Classifier` chỉ nhận Quốc hiệu khi cả đoạn đúng bằng chuỗi Quốc hiệu → cả khối đầu thành `Body`. |
| A2 | Hai khối chữ ký ("CÁN BỘ BÁO CÁO", "LÃNH ĐẠO/CHỈ HUY CÓ THẨM QUYỀN", "(Ký, ghi rõ họ tên…)", họ tên) bị kéo khỏi nửa phải: tiêu đề căn giữa cả trang, "(Ký…)" và họ tên căn đều sát lề trái. | Bản gốc đặt chữ ký bằng thụt lề trái ~10 cm; các dòng này không được nhận là khối chữ ký (thiếu "Nơi nhận", "V. Ý KIẾN…" bị nhận nhầm là chức vụ) nên `ResetSideIndents` xóa thụt lề. |
| A3 | Đề mục "II.", "III.", "IV." căn **giữa**, trong khi "I." căn đều — NĐ 30 (tr. 33, mục 6b) trình bày "I. NHỮNG KẾT QUẢ…" như đề mục thường, không căn giữa. | "II./III./IV." là đánh số tự động của Word, không có trong `Range.Text` → dòng in hoa ngắn bị nhận là *tên phần/chương* (`TenPhanChuong`). |
| A4 | "1.000.000đ" bị in đậm như đề mục (trừ NĐ 250). | `RxSoMuc` khớp "1.000." như số thứ tự "1.". |
| A5 | Ghi chú mẫu cỡ 8 ("Mẫu BL2", "BH theo TT số 44/2025/TT-BCA", "Ngày 18/6/2025", "VB21") thành chữ cỡ 14; "VB21" thành tiêu đề đậm căn giữa. | Không có loại "ghi chú/ký hiệu mẫu"; dòng in hoa ngắn → `TenPhanChuong`. |
| A6 | 4 đoạn gạch đầu dòng tự động giữ thụt −0,02 cm → **rà soát của chính V-FORMAT báo lỗi thụt dòng ngay sau khi chuẩn hóa**. | `Formatter` giữ thụt lề của đoạn trong danh sách (`inList`), `FormatChecker` vẫn đòi 1 cm. |

Ngày "tháng 09" được sửa thành "tháng 9": **đúng** NĐ 30 (chỉ ngày < 10 và tháng 1, 2 mới thêm số 0 — tr. 23).

## B. Sai lệch theo từng quy chuẩn (so với văn bản gốc)

### 1. NĐ 30/2020/NĐ-CP — Hành chính

Đầu ra trên T9 giống hệt TT 68 và Công an. Phần lớn thông số khớp Phụ lục I (lề, phông, cỡ/kiểu từng thành phần, nơi nhận, số trang).

| Quy định gốc | V-FORMAT |
|---|---|
| Số văn bản nhỏ hơn 10 phải thêm số 0 phía trước (tr. 23, mục 3c) | `FixSoKyHieu` không thêm số 0 ("Số: 5/QĐ-UBND" giữ nguyên), rà soát không báo |
| Cỡ chữ tăng, giảm phải thống nhất: Quốc hiệu 12 → Tiêu ngữ 13 → ngày tháng 13, hoặc 13 → 14 → 14 (tr. 32, chú thích 1) | Mặc định Quốc hiệu 12, Tiêu ngữ 13 nhưng ngày tháng **14** |

### 2. TT 68/2023/TT-BQP — Quân đội

**Chưa đối chiếu được** (không có toàn văn). Bộ thông số chỉ là NĐ 30 + khối đầu cỡ 13 + cấp bậc người ký, chưa có căn cứ.
Theo chính quy tắc thống nhất cỡ chữ của NĐ 30, khi Quốc hiệu cỡ 13 thì Tiêu ngữ phải cỡ 14 — profile vẫn để Tiêu ngữ 13.

### 3. VBHN 08/VBHN-BCA (2021) — Công an

Điều 9 (tr. 7) dẫn chiếu Điều 8, 9 NĐ 30 → mang cả 2 lỗi của mục 1. Cấp bậc hàm trước họ tên đúng Điều 13 (tr. 8–10).

| Quy định gốc | V-FORMAT |
|---|---|
| Đơn vị sự nghiệp được ghi chức danh khoa học, danh hiệu **trước** cấp bậc hàm, tối đa 2 (Điều 13 khoản 1 điểm d, tr. 10) | `RxRank` bắt buộc cấp bậc đứng đầu, `IsPersonName` loại tên có dấu chấm → "PGS.TS Thiếu tướng …" bị báo thiếu cấp bậc / không nhận là họ tên |
| Cấp bậc hàm gồm cả hạ sĩ quan | `RxRank` chỉ có Đại/Thượng/Trung/Thiếu + tướng/tá/úy |
| — | Báo cáo T9 ghi "BH theo TT số 44/2025/TT-BCA": có thể có văn bản mới hơn VBHN 08, chưa có trong bộ nguồn |

### 4. HD 05-HD/VPTW ngày 27/5/2026 — Đảng (thay HD 36/2018)

Đúng: tiêu đề 15 đậm có kẻ, cơ quan 14, dấu sao, "Số" không hai chấm, ngày 14 nghiêng, tên loại 15–16, trích yếu 14–15,
trích yếu công văn 12 nghiêng, nội dung 14–15, "Nơi nhận" 14 gạch chân, nơi nhận 12, T/M, họ tên 14 đậm,
số trang cỡ 14 cách mép trên 10 mm, ẩn trang đầu.

| Quy định gốc | V-FORMAT |
|---|---|
| Lề **cố định** 20 / 20 / 30 / 15 mm (tr. 28, mục IV.4) | Rà soát chấp nhận khoảng 20–25 / 30–35 / 15–20 |
| Giãn dòng **Exactly 18–22 pt** (tr. 12, mục 6.2.2) | Bội số 1,15 (người dùng chỉnh được tới 1,5) |
| **Chức vụ** người ký ("PHÓ BÍ THƯ"): in hoa 14, **đứng, không đậm** (tr. 33) | In đậm (cùng kiểu với quyền hạn ký) |
| "Kính gửi:" in thường 14 **nghiêng** (tr. 33) | Đứng |
| **Khoản** "1. …": đứng, **không đậm** (tr. 32) | Đề mục số in đậm |
| Phần, Chương, Mục và tên: 14–15 (tr. 32) | 13–14 |
| Nơi nhận kết thúc mỗi dòng bằng **dấu chấm phẩy (;)** (tr. 25, mục 9.2.5; tr. 33) | Dấu phẩy (,) |
| Dưới trích yếu có **05 dấu gạch nối "-----"** (tr. 11, mục 5.2.1) | Chuẩn hóa không chèn (chỉ có trong mẫu thư viện), rà soát không báo |
| Ký thay **K/T**, thừa lệnh **T/L** (tr. 13) | Chỉ đổi TM. → T/M; "KT.", "TL." giữ nguyên |
| Số văn bản < 10 thêm số 0 (tr. 9, mục 3.2.2) | Không thêm |
| Căn cứ: đầu dòng "-", cuối ";", căn cứ cuối "." (tr. 12, mục 6.2.3); không quy định in nghiêng | Không sửa dấu; căn cứ in nghiêng theo NĐ 30 |

### 5. HD 29-HD/TWĐTN-VP (2013) — Đoàn

Nguồn là bản đăng lại (chưa có bản ký, đóng dấu). Đúng: tiêu đề 15 đậm có kẻ, cơ quan 14, "***", "Số:", tên loại 16,
trích yếu 14 đậm, "TM.", họ tên 14 đậm, số trang.

| Quy định gốc (Mẫu 3 tr. 20–21, mục II, 5–8) | V-FORMAT |
|---|---|
| Lề phải **20 mm** | 15 mm |
| Ngày tháng cỡ **13** nghiêng | Cỡ 14, rà soát chỉ chấp nhận 14 |
| Trích yếu công văn cỡ 12 **nghiêng** | Cỡ 12 đứng |
| "Nơi nhận:" cỡ **13, đứng, đậm**, không gạch chân; danh sách cỡ **11** | Cỡ 14 gạch chân không đậm; danh sách cỡ 12 |
| Chức vụ người ký ("BÍ THƯ"): 14 **đứng, không đậm** | In đậm |
| "Kính gửi:" 14 đậm | Thường |
| Lùi đầu dòng **1,27 cm**; giãn dòng tối đa **1,2**; nội dung cỡ **14** cố định | 1 cm; cho chọn tới 1,5; cho chọn cỡ 13 |
| Dưới tên loại, trích yếu là **gạch nét rời "-------"** | Kẻ đường liền dài 0,4 dòng chữ |
| "Phần" và số thứ tự cỡ 15–16 | 13–14 |

### 6. NĐ 250/2026/NĐ-CP — Văn bản hợp nhất

| Quy định gốc (Phụ lục, tr. 3–4) | V-FORMAT |
|---|---|
| Khổ giấy, lề, phông, cỡ, kiểu chữ theo pháp luật về kỹ thuật trình bày **văn bản QPPL** | Dùng thông số NĐ 30 (hành chính) + dòng đơn; nguồn QPPL chưa có trong bộ |
| Ô trái: "VĂN BẢN HỢP NHẤT" + "Số: …" **thay cho tên cơ quan**; dưới cùng là "PHẦN KÝ XÁC THỰC" sau đường kẻ hết chiều ngang | Coi "VĂN BẢN HỢP NHẤT" là dòng tên loại thứ nhất giữa trang (test `ConsolidatedDocumentHasTwoTitleLines`); không nhận "PHẦN KÝ XÁC THỰC" |
| Số dạng `01/2026/VBHN-LQ-VPQH`: có năm, số < 10 thêm 0 | Không kiểm năm, không thêm 0 |
| Phạm vi: chỉ kỹ thuật trình bày văn bản hợp nhất | Tên "Pháp điển & hợp nhất", mã "NĐ 250" thiếu năm 2026 |

### 7. CV 5512/BGDĐT-GDTrH — Kế hoạch bài dạy

| Quy định gốc (Phụ lục IV) | V-FORMAT |
|---|---|
| Chỉ quy định **khung nội dung**: I. Mục tiêu (1–3); II. Thiết bị dạy học và học liệu; III. Tiến trình dạy học, 4 hoạt động, mỗi hoạt động a) Mục tiêu b) Nội dung c) Sản phẩm d) Tổ chức thực hiện | Rà soát **không kiểm khung** (thiếu hoạt động, thiếu a–d không báo) |
| Không quy định lề, cỡ chữ, số trang | Báo lỗi `GD5512-01..06` về lề, cỡ chữ như thể là quy định của CV 5512 (thông số do dự án tự đặt) |
| Đề mục viết thường: "I. Mục tiêu", "II. Thiết bị dạy học và học liệu"; Hoạt động 2 "Hình thành kiến thức mới/giải quyết vấn đề/thực thi nhiệm vụ đặt ra từ Hoạt động 1" | Mẫu thư viện viết hoa toàn bộ, rút gọn tên Hoạt động 2 |

## Không kiểm được

- TT 68/2023/TT-BQP: chưa có toàn văn.
- Văn bản QPPL làm nền cho NĐ 250: chưa có trong `docs/quy-chuan/`.
- Mỗi quy chuẩn mới chạy trên một văn bản (báo cáo Công an); các thành phần báo cáo T9 không có (Nơi nhận, Điều, Chương, số ký hiệu) chỉ được đối chiếu qua thông số trong `Profiles.cs`, chưa qua Word.
