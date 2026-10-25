# Kiểm thử chính tả + dấu thanh trên Word ở máy ảo (25/10/2026)

> **Cập nhật (sau khi sửa):** `TonePlacementValidator` đặt dấu theo cấu trúc âm tiết (qu/gi là phụ âm đầu; nguyên âm
> mũ/móc nhận dấu, ươ → ơ; kiểu cũ/mới chỉ khác ở oa/oe/uy không có âm cuối; bỏ qua từ dính liền/không phải âm tiết),
> `ConvertToneMarks` kiểu cũ dùng cùng quy tắc (không đảo chiều), nút **Dấu mới/Dấu cũ** gọi `Spell.ConvertTones`
> thay đúng từng từ. Unit test 227/227. Chạy lại trên Word (add-in `ACDCBD…`):
>
> | Nút | Kết quả |
> |---|---|
> | Kiểm tra chính tả | 25 lỗi, **0 báo nhầm** (hết huỵên, Nguỹên, qủa, khuýêt, Qúy, quýêt, thưchịên) |
> | Sửa tự động | Sửa 25 lỗi, "Đồng chí Nguyễn Văn Quý", "huyện Thụy Anh" giữ nguyên; giữ đậm/nghiêng; sửa cả trong bảng |
> | Dấu mới | Đổi 9 từ (hoà, Thuỷ, Toà, Thuỵ, khoẻ, bảng), 16 đoạn như cũ, **bảng nguyên vẹn**, giữ **hòa** đậm và *xuất sắc* nghiêng; Undo 1 bước |
> | Dấu cũ | Đổi Ủy → Uỷ, Thụy → Thuỵ, đi hết văn bản; hoàng/loại/khuỷu/quả không đổi; giữ định dạng, bảng |
>
> Quét 26 nghìn từ trong `samples/output`, `tmp/e2e`: chỉ đổi hòa/hoà, ủy/uỷ.
>
> **App độc lập (`DocxFormatter`):** `FixSpelling` giờ gọi `DocxWriter.FixSpelling` — sửa ngay trong các node chữ
> của run, không ghi lại cả đoạn; chỉ đoạn cần sửa chữ theo `Planner` mới ghi lại như trước. Lỗi vắt qua nhiều run:
> gợi ý cùng độ dài thì chia ký tự về đúng run cũ (đề **suất** → đề **xuất** vẫn đậm), khác độ dài thì vào run đầu.
> CLI trên file có đậm/nghiêng xen kẽ, tab, ngắt dòng: giữ **hòa**, *xuất sắc*, tab, ngắt dòng; từ đúng không đổi.
> CLI trên 36 file `samples/input` + `test/`: không lỗi, OOXML hợp lệ; mọi chỗ sửa chính tả đều đúng (khoảng trắng,
> HOÀ → HÒA, uỷ ban → Ủy ban, sử lý → xử lý). Unit test 233/233.
>
> Phần dưới là kết quả lần chạy **trước khi sửa**.

Bản thử: thay đổi chưa commit (dấu thanh trong `SpellChecker.Check`, `ConvertToneMarks`, nút **Dấu mới / Dấu cũ**,
`FixSpelling` trong `DocxFormatter`). Add-in build `40C41DAD…`, Word 16.0 (Windows 11 ARM, UTM).
Cách chạy: script tạo văn bản 9 đoạn + bảng 2×2 có lỗi thật (đoạn 6 có chữ **đậm**/*nghiêng* xen kẽ), bấm nút trên Ribbon
bằng UI Automation (nhóm "Chính tả Offline" bị thu gọn nên mở menu nhóm trước), trả lời hộp thoại, đọc lại văn bản qua COM,
lưu docx/pdf, chụp màn hình. Script: `tmp/spell-vm/` (`spell-user.ps1`, `click.ps1`). Unit test: 173/173 đạt.

## Văn bản thử

```
BÁO CÁO TÌNH HÌNH SẢN SUẤT VÀ XỬ DỤNG NGÂN SÁCH NĂM 2026
Căn cứ qui định của Ủy Ban Nhân Dân tỉnh về việc sát nhập các xã trên địa bàn;
Phòng đã bổ xung hồ sơ và đề suất phương án sử lý dứt điểm , đảm bảo năng xuất lao động.
Các đơn vị đã đã thực hiện  nghiêm chủ chương của Đảng và và Nhà nước.
Ban hoà giải cơ sở phối hợp với Chi cục Thuỷ sản và Toà án nhân dân huyện Thuỵ Anh.
Kết quả **hoà** nhập cộng đồng của học sinh khuyết tật đạt loại *xuất sắc*.
Đồng chí Nguyễn Văn Quý, ông Thụy, bà Khuê ở xóm Khuya; khuỷu tay; thuở nhỏ; quả quyết; hoàng hôn; khoẻ mạnh; ...
Xin trân thành cảm ơn các đơn vị đã sắp sếp thời gian, không dấu diếm khuyết điểm và rút ra bài học xâu xắc.
Lỗi gõ máy: Vieejt Nam, chungs tooi, nhân dânn, thựchiện, kế hoach, ngân sach, viêc làm.
| Chi cục Thuỷ sản | sử lý hồ sơ tồn đọng |
| Tổ hoà giải      | Toà án nhân dân huyện |
```

## Kết quả theo nút

| Nút | Kết quả |
|---|---|
| **Kiểm tra chính tả** | Hộp thoại "Phát hiện 35 lỗi", 35 nhận xét + tô vàng; văn bản không đổi, Undo được. **25 đúng**: sản suất, xử dụng, qui định, Ủy Ban Nhân Dân, sát nhập, bổ xung, đề suất, sử lý (cả trong bảng), năng xuất, chủ chương, trân thành, đã đã, và và, 2 khoảng trắng, khoảng trắng trước dấu phẩy, hoà/Thuỷ/Toà/Thuỵ/khoẻ. **10 sai (báo nhầm từ đúng)**: huyện→*huỵên* (×2), quả→*qủa* (×2), khuyết→*khuýêt* (×2), Nguyễn→*Nguỹên*, Quý→*Qúy*, quyết→*quýêt*, thựchiện→*thưchịên*. |
| **Sửa tự động** | Hỏi xác nhận → "Đã sửa 35 lỗi". Giữ được chữ đậm/nghiêng (chỉ thay đúng đoạn lỗi), sửa cả trong bảng. Nhưng **ghi luôn 10 gợi ý sai vào văn bản**: "Đồng chí Nguỹên Văn Qúy", "huỵên Thụy Anh", "Kết qủa", "khuýêt tật". |
| **Dấu mới (hòa, thủy)** | Đổi đúng hoà/Thuỷ/Toà/khoẻ, nhưng: (1) cùng 10 từ đúng bị làm hỏng như trên; (2) **vỡ bảng**: đoạn "Lỗi gõ máy…" bị kéo vào ô đầu của bảng, các ô có thêm dấu xuống dòng, số đoạn 16 → 19; (3) **mất định dạng trong đoạn**: "**hoà**" đậm và "*xuất sắc*" nghiêng ở đoạn 6 thành chữ thường. Undo 1 bước khôi phục được. |
| **Dấu cũ (hoà, thuỷ)** | Chỉ đổi đoạn đầu tiên có thay đổi (Ủy → Uỷ) rồi dừng; các đoạn sau không được xử lý (chạy 2 lần, cùng kết quả). Trong lúc chạy, đọc qua COM thấy đoạn 2 dính văn bản đoạn 3. |

App độc lập (`DocxFormatter`, `FixSpelling` **mặc định bật**), chạy CLI trên cùng file: cũng ghi "Nguỹên Văn Qúy", "huỵên",
"qủa", "khuýêt" vào văn bản, và đoạn nào có sửa chính tả thì mất chữ đậm/nghiêng xen kẽ (`DocxWriter.SetText` ghi lại cả đoạn).

## Nguyên nhân

1. `TonePlacementValidator.IsTonePlacedProperly` — trước đây chưa dùng, nay được gọi trong `Check`:
   - coi `u` của **qu** và `i` của **gi** là nguyên âm → quả→qủa, Quý→Qúy;
   - âm tiết 3 nguyên âm luôn đặt dấu vào nguyên âm giữa → **uyê** (huyện, khuyết, quyết, Nguyễn) sai; quy tắc đúng: có
     nguyên âm mũ/móc (â ê ô ơ ư, ă) thì dấu đặt ở đó (ươ → ơ);
   - không xét chữ dính liền (thựchiện).
2. `ConvertToneMarks(newStyle: false)` dùng `GetAlternativeTonePlacement`, hàm này **đảo chiều cả hai phía** (hoà → hòa) và
   không xét âm cuối/âm tiết 3 nguyên âm: hoàng → *hòang*, loại → *lọai*, khuỷu → *khủyu* (thấy khi chạy thử trên máy, chưa thấy trong
   Word vì vòng lặp đã dừng sớm).
3. Nút **Dấu mới/Dấu cũ** gán `p.Range.Text = …` cho **cả đoạn** (kèm dấu đoạn `\r` / dấu ô `\r\a`) trong lúc đang duyệt
   `d.Paragraphs` → vỡ bảng, mất định dạng ký tự, vòng lặp hỏng. Cần thay đúng từng từ (như `Spell.Run` đang làm) và không
   đụng dấu đoạn/ô.

## Chưa bắt được (ngoài từ điển, đúng như thiết kế hiện tại)

sắp sếp, dấu diếm, xâu xắc, và lỗi gõ Telex (Vieejt, chungs tooi, dânn, hoach, sach, viêc). `SyllableValidator` mới thêm
có thể bắt một phần (âm tiết không hợp lệ) nhưng chưa được nối vào `SpellChecker`.

## Ghi chú nhỏ

- Nhận xét cho lỗi khoảng trắng trước dấu câu hiện "nên viết """ — khó hiểu với người dùng.
