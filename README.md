# V-FORMAT — Chuẩn hóa thể thức văn bản Word 1-Click

Add-in tích hợp trực tiếp vào Microsoft Word trên Windows (Office 2013+ / Microsoft 365, 32 và 64 bit).
Chạy **100% offline**, không gửi dữ liệu văn bản ra ngoài.

Có hai cách dùng:

1. **App V-FORMAT (`VFormat.exe`)** — chạy độc lập, không cần mở Word: kéo thả tệp `.docx` (hoặc cả thư mục) vào cửa sổ,
   app tạo tệp `<tên>_chuan.docx` đúng thể thức; tab *Tạo văn bản từ mẫu* sinh văn bản mới từ thư viện mẫu.
2. **Add-in trong Microsoft Word** — tab V-FORMAT trên Ribbon (chuẩn hóa, rà soát, chính tả, chèn thành phần...).

## App V-FORMAT (kéo thả)

- Tải `VFormat-App.zip`, giải nén, chạy `VFormat.exe` (Windows 10/11 64 bit; không cần cài .NET, không cần Word).
  Windows SmartScreen có thể cảnh báo vì tệp chưa ký số: *More info → Run anyway*.
- Chọn quy chuẩn (mặc định **Tự nhận diện**), cỡ chữ 13/14, giãn dòng, nơi lưu (cạnh tệp gốc hoặc thư mục khác).
- Kéo thả tệp/thư mục vào khung, hoặc bấm vào khung để chọn nhiều tệp. Kéo thả tệp lên biểu tượng `VFormat.exe` cũng được.
- Bảng kết quả: quy chuẩn nhận diện được, số lỗi thể thức trước → sau; bấm đúp để mở tệp, *Xem báo cáo* để xem chi tiết.
- Chỉ nhận `.docx`; tệp `.doc` cũ hãy mở bằng Word và *Save As* `.docx` trước. Tệp gốc không bị sửa.

## Hỗ trợ 7 hệ quy chuẩn

| Mã | Văn bản quy định | Phạm vi |
|---|---|---|
| NĐ 30 | Nghị định 30/2020/NĐ-CP | Hành chính Nhà nước |
| TT 68 | Thông tư 68/2023/TT-BQP | Văn bản Quân đội |
| VBHN 08 | Văn bản hợp nhất 08/VBHN-BCA (2021) | Văn bản Công an nhân dân |
| HD 05 | Hướng dẫn 05-HD/VPTW | Văn bản Đảng |
| HD 29 | Hướng dẫn 29-HD/TWĐTN-VP | Văn bản Đoàn TNCS Hồ Chí Minh |
| NĐ 250/2026 | Nghị định 250/2026/NĐ-CP (Phụ lục: kỹ thuật trình bày văn bản hợp nhất) | Văn bản hợp nhất |
| CV 5512 | Công văn 5512/BGDĐT-GDTrH (Phụ lục IV) | Kế hoạch bài dạy |

TT 68/2023/TT-BQP: **chưa có toàn văn công khai để kiểm chứng** (xem `docs/quy-chuan/02-TT68-2023-BQP-CHUA-TIM-THAY.md`);
thông số Quân đội hiện là NĐ 30 + khối đầu 13/14/14 + cấp bậc quân hàm người ký.

Văn bản hợp nhất (NĐ 250/2026, hiệu lực 01/7/2026): ô trái khối đầu ghi **"VĂN BẢN HỢP NHẤT" + `Số: 01/2026/VBHN-...`** thay cho
tên cơ quan (dòng "VĂN BẢN HỢP NHẤT" đặt giữa trang được chuyển vào ô trái); tên loại giữa trang là tên loại của văn bản được
sửa đổi (THÔNG TƯ, NGHỊ ĐỊNH...). Phần ký xác thực ngăn với nội dung bằng **đường kẻ ngang dài nét liền**; cột phải: tên cơ quan
ký xác thực (kẻ dưới) → "XÁC THỰC VĂN BẢN HỢP NHẤT" → ngày tháng → chức vụ → (Chữ ký và dấu) → họ tên. Rà soát báo thiếu
"VĂN BẢN HỢP NHẤT", số thiếu năm/ký hiệu VBHN, thiếu phần ký xác thực. Khổ giấy, lề, phông, cỡ chữ: Phụ lục dẫn chiếu quy định
về văn bản QPPL — V-Format dùng thông số NĐ 30, dòng đơn.

Kế hoạch bài dạy (CV 5512 Phụ lục IV): rà soát **khung nội dung** (mã `GD5512-11`) — Tên bài dạy, Môn học/lớp, Thời gian
thực hiện; I. Mục tiêu (1. Về kiến thức, 2. Về năng lực, 3. Về phẩm chất); II. Thiết bị dạy học và học liệu; III. Tiến trình
dạy học với 4 hoạt động, mỗi hoạt động có a) Mục tiêu, b) Nội dung, c) Sản phẩm, d) Tổ chức thực hiện (thiếu Hoạt động 4 Vận
dụng chỉ nhắc, theo Ghi chú 1). Công văn không quy định thể thức trình bày, nên lỗi lề, phông, cỡ chữ của KHBD chỉ ở mức
**khuyến nghị** (Info).

Công an (08/VBHN-BCA, hợp nhất Thông tư 26/2019/TT-BCA về công tác văn thư trong CAND): Điều 9 dẫn chiếu thể thức theo
Điều 8, 9 NĐ 30/2020, nên V-Format trình bày như NĐ 30; phần riêng của CAND được rà soát là **cấp bậc hàm trước họ tên người ký**
(ví dụ "Đại tá Nguyễn Văn A"; nhận cả hạ sĩ quan "Thượng sĩ/Trung sĩ/Hạ sĩ", chiến sĩ "Binh nhất/Binh nhì", và chức danh khoa học,
danh hiệu đứng trước cấp bậc như "PGS.TS Thiếu tướng ..." theo Điều 13) — app báo lỗi `CA08-09` nhưng không tự điền vì không biết cấp bậc.

Văn bản thực tế thường gặp cũng được xử lý: khối đầu gõ hai cột bằng **tab** (tách về bảng 2 cột, nối tên cơ quan bị xuống dòng);
**ghi chú biểu mẫu** ("Mẫu BL2", "BH theo TT số …", "VB21") giữ nguyên; **khối chữ ký đặt bằng thụt lề** giữ ở nửa phải;
đề mục **đánh số tự động** của Word ("II.", "III.") nhận theo số Word hiển thị; danh sách tự động không bị báo sai thụt dòng.
Văn bản Quân đội (TT 68) cũng được rà soát cấp bậc quân hàm người ký như vậy (mã `QD68-09`).

NĐ 30: số văn bản nhỏ hơn 10 thêm số 0 (`Số: 05/QĐ-UBND`); cỡ chữ khối đầu thống nhất — Quốc hiệu 12, Tiêu ngữ 13, ngày tháng 13
(Quân đội: 13/14/14), rà soát báo lỗi khi lệch.

Đảng (HD 05-HD/VPTW ngày 27/5/2026): lề cố định 20/20/30/15 mm; nội dung, căn cứ, Điều, đề mục, "Kính gửi" cỡ 14–15
(chọn cỡ 13 vẫn giữ 14), giãn dòng Exactly 18 pt; trích yếu công văn cỡ 12 nghiêng; dưới trích yếu chèn `-----`;
khoản không đậm; căn cứ `- Căn cứ ...;` (căn cứ cuối dấu chấm), kiểu chữ đứng; "Kính gửi:" nghiêng; nơi nhận kết thúc bằng `;`;
`T/M`, `K/T`, `T/L`, `T/UQ`; chức vụ người ký ("PHÓ BÍ THƯ") không đậm.

Đoàn (HD 29-HD/TWĐTN-VP, Mẫu 3): lề 20/20/30/20 mm; lời văn cỡ 14 cố định, lùi đầu dòng 1,27 cm, giãn dòng tối đa 1,2;
ngày tháng cỡ 13; trích yếu công văn cỡ 12 nghiêng; dưới trích yếu chèn `-------`; "Kính gửi:" đậm (chữ "Kính gửi:" nghiêng);
"Nơi nhận:" cỡ 13 đậm, danh sách cỡ 11; chức vụ người ký không đậm.

Mặc định **tự nhận diện** quy chuẩn (theo tiêu đề Đảng/Đoàn, tên cơ quan quân đội, "KẾ HOẠCH BÀI DẠY", "VĂN BẢN HỢP NHẤT"...),
hoặc chọn cố định trên Ribbon / bảng điều khiển.

## Chức năng (tab **V-FORMAT** trên Ribbon)

- **Chuẩn hóa 1-Click**: khổ A4 + lề theo quy chuẩn; phông Times New Roman; cỡ/kiểu chữ từng thành phần;
  sửa chữ Quốc hiệu, Tiêu ngữ, số ký hiệu (`Số: 12/QĐ-UBND`, Đảng `Số 12-NQ/TU`), ngày tháng
  (thêm số 0: ngày < 10, tháng 1–2), "Kính gửi:", "Điều 1.", danh sách nơi nhận (`- ...;` … `- Lưu: VT.`), quyền hạn ký (`TM.`/`KT.`/`TL.`, Đảng `T/M`/`K/T`/`T/L`);
  thụt đầu dòng 1 cm, giãn dòng, khoảng cách đoạn; tự **dàn khối đầu và khối chữ ký – nơi nhận thành bảng 2 cột không viền**;
  **kẻ đường** dưới tiêu ngữ (bằng độ dài dòng chữ), dưới tên cơ quan ban hành và trích yếu (≈ 1/3–1/2; Đảng, Đoàn dùng dòng gạch `-----`);
  **số trang** giữa đầu trang, ẩn ở trang đầu; bảng nội dung vừa khổ giấy và tự lặp hàng tiêu đề. Tất cả gom thành **một bước Undo**.
- **Rà soát thể thức**: phát hiện lỗi và gắn nhận xét "V-Format" vào đúng đoạn. Mã lỗi `<quy chuẩn>-01..10`
  (vd. `ND30-04`, `DTN-07`): 01 khổ giấy · 02 lề · 03 phông · 04 cỡ chữ · 05 đậm/nghiêng/in hoa · 06 căn lề/thụt dòng ·
  07 Quốc hiệu–tiêu ngữ–tiêu đề · 08 số ký hiệu/ngày tháng · 09 chữ ký/nơi nhận · 10 thiếu thành phần/khoảng trắng/số trang ·
  11 khung nội dung kế hoạch bài dạy (CV 5512).
  Nút **Xóa nhận xét V-Format** / **Xóa tất cả nhận xét**.
- **Bố cục trang in A4**: lề trang, đánh số trang, **in một mặt / hai mặt** (lề đối xứng, nhớ riêng cho từng tài liệu), chuẩn hóa bảng.
- **Chèn thành phần thể thức**: khối đầu (theo quy chuẩn đang dùng), khối chữ ký & nơi nhận, nơi nhận, địa danh – ngày tháng hôm nay.
- **Chính tả Offline**: lỗi phụ âm/vần hay gặp (sử lý → xử lý, bổ xung → bổ sung…), viết hoa tên cơ quan theo Phụ lục II NĐ 30
  (uỷ ban nhân dân → Ủy ban nhân dân…), hư từ lặp, khoảng trắng. Tô vàng + nhận xét, hoặc **Sửa tự động**.
  **Từ điển riêng**: `%APPDATA%\VFormat\tu-dien.txt`, mỗi dòng `sai=đúng`.
- **Thư viện mẫu**: 27 mẫu dựng sẵn (Hành chính, Quân đội, Công an, Đảng, Đoàn, Pháp điển, Giáo dục) và **kho mẫu cá nhân**
  (`%APPDATA%\VFormat\Templates`, nút *Lưu làm mẫu*).
- **Bảng điều khiển** (Task Pane bên phải): chọn quy chuẩn, cỡ chữ 13/14, giãn dòng, địa danh, bật/tắt từng bước chuẩn hóa,
  danh sách lỗi — bấm vào dòng để nhảy tới đoạn văn.

## Cài đặt (trên máy Windows)

1. Cần .NET Framework 4.8 (có sẵn trên Windows 10/11) và Microsoft Word 2013 trở lên.
2. Giải nén `VFormat.zip`, **đóng Word**, chạy `install.cmd` (không cần quyền Administrator).
3. Mở Word → tab **V-FORMAT**.

Gỡ cài đặt: chạy `uninstall.cmd`. Nếu tab không hiện: *File → Options → Add-ins → Manage: COM Add-ins → Go…* và tick **V-FORMAT**.

## Build

```bash
./build.sh            # chạy test Core, build Release, tạo dist/VFormat.zip
```

Yêu cầu .NET SDK 6+ (build được cả trên macOS/Linux nhờ reference assemblies; kiểu interop Office được nhúng – không cần PIA).
Bộ cài `.exe` tùy chọn: `iscc installer\VFormat.iss` (Inno Setup 6, trên Windows).

## Trạng thái kiểm thử

- Lõi `VFormat.Core`: unit test pass (`dotnet test`).
- Bộ mẫu `samples/`: **đầu vào** do `tools/make_samples.py` sinh tự động (cố ý sai thể thức: Arial, khổ Letter, sai Quốc hiệu,
  tiêu ngữ, số ký hiệu, ngày tháng, nơi nhận...); **đầu ra** do `tools/VFormat.Cli` tạo hoàn toàn tự động bằng cùng lõi với add-in,
  không chỉnh tay. Đầu ra hợp lệ schema OOXML; mẫu 01–05 đã mở kiểm tra bằng Microsoft Word, mẫu 06 (Công an) mới kiểm qua ảnh Quick Look.
- App `VFormat.exe`: build/publish thành công (PE x64), engine đã kiểm với 6 mẫu + 27 mẫu thư viện (schema hợp lệ; còn 0 lỗi, trừ mẫu 06 cố ý thiếu cấp bậc hàm người ký);
  **giao diện chưa chạy thử trên Windows**.
- Add-in COM trên Windows (Ribbon, Task Pane, vẽ đường kẻ bằng shape): build thành công, **chưa kiểm thử trên Word for Windows**.

## Test không cần Word (macOS/Linux/Windows)

`tools/VFormat.Cli` chạy cùng lõi `VFormat.Core` với add-in trên file .docx (OpenXML), kèm báo cáo nhận diện + lỗi trước/sau
và kiểm tra hợp lệ schema OOXML:

```bash
python3 tools/make_samples.py                      # tạo văn bản sai thể thức vào samples/input
dotnet run --project tools/VFormat.Cli -- samples/input/01-quyet-dinh-hanh-chinh.docx out.docx --report out.md
python3 tools/compare_page.py                      # samples/compare.html: ảnh Trước | Sau + báo cáo
```

Khác biệt với add-in: CLI kẻ đường bằng viền dưới của một đoạn mỏng (độ dài ước lượng theo bảng độ rộng ký tự Times New Roman),
add-in đo vị trí chữ thật trong Word và vẽ shape.

## Cấu trúc

```
src/VFormat.Core/        Lõi thuần .NET Standard 2.0 (không phụ thuộc Word, có unit test)
  Profiles.cs            Thông số 7 quy chuẩn (lề, cỡ chữ, kiểu chữ từng thành phần)
  Classifier.cs          Nhận diện thành phần thể thức từng đoạn
  StandardDetector.cs    Tự nhận diện quy chuẩn
  TextNormalizer.cs      Sửa chữ Quốc hiệu, số ký hiệu, ngày tháng, nơi nhận...
  FormatChecker.cs       Rà soát, sinh mã lỗi
  SpellChecker.cs        Chính tả offline
  TemplateLibrary.cs     Thư viện mẫu
  Planner.cs             Kế hoạch dàn bảng / định dạng từng đoạn (dùng chung add-in + CLI)
src/VFormat.AddIn/       COM add-in Word (net48): Connect (Ribbon), TaskPaneHost, Formatter, Layout, PageNumbers, Services
tests/VFormat.Core.Tests Unit test lõi
src/VFormat.Docx/        Engine chuẩn hóa .docx bằng OpenXML (không cần Word) — dùng cho app và CLI
src/VFormat.App/         App WinForms kéo thả (net10.0-windows, publish một tệp VFormat.exe)
tools/VFormat.Cli        Chuẩn hóa .docx bằng dòng lệnh (test không cần Word)
samples/                 Văn bản mẫu: input (sai thể thức), output, render, compare.html
installer/               install.ps1 / uninstall.ps1 (đăng ký COM per-user), VFormat.iss
```

Thông số mặc định của từng quy chuẩn nằm trong `Profiles.cs`; khi cơ quan có quy định riêng, chỉnh tại đây rồi build lại.
