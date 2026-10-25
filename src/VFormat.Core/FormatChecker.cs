using System;
using System.Collections.Generic;
using System.Linq;

namespace VFormat.Core
{
    /// <summary>
    /// Rà soát thể thức. Mã lỗi gồm tiền tố quy chuẩn + nhóm 01..10:
    /// 01 khổ giấy · 02 lề trang · 03 phông chữ · 04 cỡ chữ · 05 kiểu chữ (đậm/nghiêng/in hoa)
    /// 06 căn lề/thụt dòng · 07 Quốc hiệu, tiêu ngữ, tiêu đề · 08 số ký hiệu, ngày tháng
    /// 09 chữ ký, nơi nhận · 10 thiếu thành phần, khoảng trắng, số trang · 11 khung nội dung (kế hoạch bài dạy CV 5512).
    /// </summary>
    public static class FormatChecker
    {
        public const int PaperGroup = 1, MarginGroup = 2, FontGroup = 3, SizeGroup = 4, StyleGroup = 5,
            AlignGroup = 6, HeaderGroup = 7, NumberDateGroup = 8, SignGroup = 9, MiscGroup = 10, StructureGroup = 11;

        public static string Code(FormatProfile p, int group) => $"{p.IssuePrefix}-{group:00}";

        public static List<Issue> Check(ClassifiedDocument doc, FormatProfile p, PageInfo page = null)
        {
            // số thập phân trong thông báo luôn theo định dạng Việt Nam ("31,8 mm") bất kể ngôn ngữ của Windows
            var saved = System.Globalization.CultureInfo.CurrentCulture;
            System.Globalization.CultureInfo.CurrentCulture = TextUtil.Vi;
            try { return CheckCore(doc, p, page); }
            finally { System.Globalization.CultureInfo.CurrentCulture = saved; }
        }

        static List<Issue> CheckCore(ClassifiedDocument doc, FormatProfile p, PageInfo page)
        {
            var issues = new List<Issue>();
            void Add(int group, int para, string msg, Severity sev = Severity.Warning, bool fix = true)
            {
                // CV 5512 không quy định thể thức trình bày: lỗi lề, phông, cỡ, kiểu chữ chỉ là khuyến nghị
                if (p.FormatAdvisory && group != StructureGroup)
                {
                    sev = Severity.Info;
                    msg += $" (khuyến nghị — {p.FullName} không quy định thể thức trình bày)";
                }
                issues.Add(new Issue { Code = Code(p, group), ParagraphIndex = para, Message = msg, Severity = sev, AutoFixable = fix });
            }

            if (page != null) CheckPage(page, p, Add, Planner.HasContentTable(doc));

            // Thiếu thành phần bắt buộc
            bool isKhbd = p.Kind == StandardKind.GiaoDuc_CV5512;
            if (!isKhbd)
            {
                if (p.HeaderTitle == null && !doc.Has(ElementKind.QuocHieu))
                    Add(MiscGroup, -1, "Thiếu Quốc hiệu \"CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM\".", Severity.Error, false);
                if (p.HeaderTitle == null && doc.Has(ElementKind.QuocHieu) && !doc.Has(ElementKind.TieuNgu))
                    Add(MiscGroup, -1, "Thiếu Tiêu ngữ \"Độc lập - Tự do - Hạnh phúc\".", Severity.Error, false);
                if (p.Kind == StandardKind.Dang_HD05 && !doc.Has(ElementKind.TieuDeDang))
                    Add(MiscGroup, -1, "Thiếu tiêu đề \"ĐẢNG CỘNG SẢN VIỆT NAM\".", Severity.Error, false);
                if (p.Kind == StandardKind.Doan_HD29 && !doc.Has(ElementKind.TieuDeDoan))
                    Add(MiscGroup, -1, "Thiếu tiêu đề \"ĐOÀN TNCS HỒ CHÍ MINH\".", Severity.Error, false);
                if (!doc.Has(ElementKind.SoKyHieu))
                    Add(MiscGroup, -1, "Thiếu số, ký hiệu văn bản.", Severity.Warning, false);
                if (!doc.Has(ElementKind.DiaDanhNgayThang))
                    Add(MiscGroup, -1, "Thiếu địa danh và thời gian ban hành.", Severity.Warning, false);
                if (!doc.Has(ElementKind.ChucVuKy))
                    Add(SignGroup, -1, "Không tìm thấy quyền hạn, chức vụ người ký.", Severity.Warning, false);
                if (!doc.Has(ElementKind.NoiNhanTitle))
                    Add(SignGroup, -1, "Thiếu phần \"Nơi nhận:\".", Severity.Warning, false);
                int dashesAfter = Planner.PlanTitleDashes(doc, p);
                if (dashesAfter >= 0)
                    Add(HeaderGroup, dashesAfter, $"Thiếu dòng gạch \"{p.TitleDashes}\" dưới tên loại, trích yếu văn bản theo {p.FullName}.");
                if (p.HeaderTitle == null) CheckHeaderSizes(doc, Add);
                int stray = Planner.StrayCongVanLabel(doc);
                if (stray >= 0) Add(MiscGroup, stray, "Công văn không ghi tên loại — bỏ dòng \"công văn\".");
                var (subject, _) = Planner.PlanCongVanSubjectMove(doc);
                if (subject >= 0) Add(NumberDateGroup, subject, "Trích yếu công văn phải đặt dưới số, ký hiệu văn bản (trong khối đầu).");
                foreach (var i in doc.Splits.Keys)
                    Add(HeaderGroup, i, "Khối đầu gõ hai cột bằng tab — phải trình bày thành bảng 2 cột (tên cơ quan bên trái, Quốc hiệu bên phải).");
                if (p.Kind == StandardKind.PhapDien_ND250) CheckConsolidated(doc, p, Add);
            }
            else CheckLessonPlan(doc, Add);

            for (int i = 0; i < doc.Paragraphs.Count; i++)
            {
                var kind = doc.Kinds[i];
                var para = doc.Paragraphs[i];
                if (kind == ElementKind.Empty || kind == ElementKind.GhiChuMau) continue;   // ghi chú biểu mẫu giữ nguyên
                var st = p.StyleFor(kind);
                var text = TextUtil.Clean(para.Text);

                if (para.FontName != null && !string.Equals(para.FontName, p.FontName, StringComparison.OrdinalIgnoreCase))
                    Add(FontGroup, i, $"Phông chữ \"{para.FontName}\" — phải dùng {p.FontName}.");
                // dòng khối đầu gõ tab (đã báo ở trên) và "(Ký, ghi rõ họ tên)": không kiểm cỡ, kiểu chữ từng thành phần
                if (doc.Splits.ContainsKey(i) || kind == ElementKind.GhiChuKy) continue;

                if (kind != ElementKind.TableText && para.FontSize > 0 && (para.FontSize < st.MinSize - 0.01 || para.FontSize > st.MaxSize + 0.01))
                    Add(SizeGroup, i, $"{Describe(kind)}: cỡ chữ {para.FontSize:0.#} — quy định {Range(st)}.");

                if (para.Colored)
                    Add(StyleGroup, i, $"{Describe(kind)}: chữ có màu — phải màu đen.");
                if (kind == ElementKind.TableText && (para.Bold == 1 || para.Italic == 1) && para.CellRow > 1)
                    Add(StyleGroup, i, "Nội dung bảng in đậm / nghiêng cả ô — chỉ hàng tiêu đề in đậm.");
                if (kind == ElementKind.Body)
                {
                    if (para.Bold == 1) Add(StyleGroup, i, "Nội dung in đậm cả đoạn — lời văn kiểu chữ đứng.");
                    if (para.Italic == 1) Add(StyleGroup, i, "Nội dung in nghiêng cả đoạn — lời văn kiểu chữ đứng.");
                }
                if (kind != ElementKind.TableText && kind != ElementKind.Body)
                {
                    bool wantBold = Planner.WantBold(kind, text, p);
                    if (para.Bold >= 0 && (para.Bold == 1) != wantBold)
                        Add(StyleGroup, i, $"{Describe(kind)} phải {(wantBold ? "in đậm" : "không in đậm")}.");
                    bool mixedItalic = kind == ElementKind.KinhGui && p.KinhGuiPrefixItalic;   // chỉ "Kính gửi:" nghiêng
                    if (para.Italic >= 0 && !mixedItalic && (para.Italic == 1) != st.Italic)
                        Add(StyleGroup, i, $"{Describe(kind)} phải {(st.Italic ? "in nghiêng" : "kiểu chữ đứng")}.");
                }
                if (kind == ElementKind.HoTenKy && p.RankBeforeSignerName && !Classifier.HasRank(text))
                    Add(SignGroup, i, $"Ghi cấp bậc hàm trước họ tên người ký (ví dụ: \"Đại tá {text}\") theo {p.FullName}.",
                        Severity.Warning, false);
                if (st.UpperCase && !TextUtil.IsAllUpper(text))
                    Add(StyleGroup, i, $"{Describe(kind)} phải viết in hoa.");

                if (para.Alignment.HasValue && !para.InTable && kind != ElementKind.TableText && para.Alignment.Value != st.Align
                    && !(st.Align == Align.Justify && para.Alignment == Align.Left && text.Length < 60))
                    Add(AlignGroup, i, $"{Describe(kind)}: căn lề {Name(para.Alignment.Value)} — quy định {Name(st.Align)}.");

                if (!para.InTable && !para.InList && st.FirstIndentCm > 0 && Math.Abs(para.FirstIndentPt - st.FirstIndentCm * 28.3465f) > 3
                    && (kind == ElementKind.Body || kind == ElementKind.Dieu || kind == ElementKind.CanCu))
                    Add(AlignGroup, i, $"Thụt đầu dòng {para.FirstIndentPt / 28.3465f:0.##} cm — quy định {st.FirstIndentCm:0.##} cm.");

                // Nội dung chữ
                var fixedText = Planner.FixText(doc, i, p, text);
                if (TextUtil.Clean(fixedText) != text)
                {
                    int group;
                    switch (kind)
                    {
                        case ElementKind.QuocHieu: case ElementKind.TieuNgu: case ElementKind.TieuDeDang: case ElementKind.TieuDeDoan:
                        case ElementKind.DauSao:
                            group = HeaderGroup; break;
                        case ElementKind.SoKyHieu: case ElementKind.DiaDanhNgayThang: case ElementKind.TrichYeuCongVan:
                            group = NumberDateGroup; break;
                        case ElementKind.NoiNhanTitle: case ElementKind.NoiNhanItem: case ElementKind.ChucVuKy:
                            group = SignGroup; break;
                        default: group = MiscGroup; break;
                    }
                    var msg = group == MiscGroup
                        ? "Khoảng trắng / dấu câu chưa chuẩn."
                        : $"{Describe(kind)} chưa đúng: \"{Trim(text)}\" → \"{Trim(TextUtil.Clean(fixedText))}\".";
                    Add(group, i, msg, group == MiscGroup ? Severity.Info : Severity.Warning);
                }
            }
            return issues;
        }

        /// <summary>
        /// NĐ 30 Phụ lục I mục V (chú thích): cỡ chữ khối đầu tăng, giảm thống nhất — Quốc hiệu 12, Tiêu ngữ 13, ngày tháng 13
        /// hoặc Quốc hiệu 13, Tiêu ngữ 14, ngày tháng 14.
        /// </summary>
        static void CheckHeaderSizes(ClassifiedDocument doc, Action<int, int, string, Severity, bool> add)
        {
            int qh = doc.Kinds.IndexOf(ElementKind.QuocHieu), tn = doc.Kinds.IndexOf(ElementKind.TieuNgu),
                nt = doc.Kinds.IndexOf(ElementKind.DiaDanhNgayThang);
            float Size(int i) => i >= 0 && !doc.Splits.ContainsKey(i) ? doc.Paragraphs[i].FontSize : 0;
            float sq = Size(qh), st = Size(tn), sn = Size(nt);
            if (sq > 0 && st > 0 && Math.Abs(st - sq - 1) > 0.01)
                add(HeaderGroup, tn, $"Cỡ chữ khối đầu chưa thống nhất: Quốc hiệu {sq:0.#}, Tiêu ngữ {st:0.#} — Tiêu ngữ lớn hơn Quốc hiệu 1 cỡ (12/13 hoặc 13/14).",
                    Severity.Warning, true);
            if (st > 0 && sn > 0 && Math.Abs(sn - st) > 0.01)
                add(HeaderGroup, nt, $"Cỡ chữ khối đầu chưa thống nhất: ngày tháng cỡ {sn:0.#} — phải bằng Tiêu ngữ (cỡ {st:0.#}).",
                    Severity.Warning, true);
        }

        static readonly System.Text.RegularExpressions.Regex RxVbhnNumber =
            new System.Text.RegularExpressions.Regex(@"^Số:?\s*[\d\.…_ ]*/\d{4}/VBHN-[A-ZĐ]");

        /// <summary>NĐ 250/2026/NĐ-CP, Phụ lục: khối đầu, số ký hiệu và phần ký xác thực của văn bản hợp nhất.</summary>
        static void CheckConsolidated(ClassifiedDocument doc, FormatProfile p, Action<int, int, string, Severity, bool> add)
        {
            bool labelInHeader = Enumerable.Range(0, doc.Kinds.Count).Any(i =>
                (doc.Kinds[i] == ElementKind.CoQuanBanHanh || doc.Kinds[i] == ElementKind.CoQuanChuQuan)
                    && TextUtil.Key(doc.Paragraphs[i].Text) == "VAN BAN HOP NHAT")
                || doc.Splits.Values.Any(x => TextUtil.Key(x.Left) == "VAN BAN HOP NHAT");
            int outside = Planner.ConsolidatedLabelOutsideHeader(doc, p);
            if (outside >= 0)
                add(HeaderGroup, outside, "\"VĂN BẢN HỢP NHẤT\" phải đặt ở ô trái khối đầu, ngay trên số, ký hiệu (không phải tên loại giữa trang).",
                    Severity.Warning, true);
            else if (!labelInHeader)
                add(HeaderGroup, -1, "Thiếu dòng \"VĂN BẢN HỢP NHẤT\" ở ô trái khối đầu (thay cho tên cơ quan).", Severity.Error, false);
            else if (doc.Has(ElementKind.CoQuanChuQuan))
                add(HeaderGroup, doc.Kinds.IndexOf(ElementKind.CoQuanChuQuan),
                    "Văn bản hợp nhất: ô trái chỉ ghi \"VĂN BẢN HỢP NHẤT\" và số, ký hiệu — không ghi tên cơ quan.", Severity.Warning, false);

            int so = doc.Kinds.IndexOf(ElementKind.SoKyHieu);
            string soText = so >= 0 ? TextUtil.Clean(doc.Paragraphs[so].Text)
                : doc.Splits.Values.Where(x => x.LeftKind == ElementKind.SoKyHieu).Select(x => x.Left).FirstOrDefault();
            if (soText != null && !RxVbhnNumber.IsMatch(soText))
                add(NumberDateGroup, so, $"Số, ký hiệu văn bản hợp nhất phải có năm ký xác thực và ký hiệu VBHN (ví dụ \"Số: 01/2026/VBHN-BTP\") — đang là \"{soText}\".",
                    Severity.Warning, false);

            if (!doc.Has(ElementKind.XacThuc))
                add(SignGroup, -1, "Thiếu phần ký xác thực \"XÁC THỰC VĂN BẢN HỢP NHẤT\".", Severity.Error, false);
            else if (!doc.Has(ElementKind.CoQuanXacThuc))
                add(SignGroup, doc.Kinds.IndexOf(ElementKind.XacThuc),
                    "Thiếu tên cơ quan của người có thẩm quyền ký xác thực (ngay trên dòng \"XÁC THỰC VĂN BẢN HỢP NHẤT\").", Severity.Warning, false);
        }

        static readonly string[] Steps = { "MUC TIEU", "NOI DUNG", "SAN PHAM", "TO CHUC THUC HIEN" };

        /// <summary>
        /// CV 5512/BGDĐT-GDTrH Phụ lục IV — khung kế hoạch bài dạy: I. Mục tiêu (kiến thức, năng lực, phẩm chất);
        /// II. Thiết bị dạy học và học liệu; III. Tiến trình dạy học gồm 4 hoạt động, mỗi hoạt động có
        /// a) Mục tiêu, b) Nội dung, c) Sản phẩm, d) Tổ chức thực hiện.
        /// </summary>
        static void CheckLessonPlan(ClassifiedDocument doc, Action<int, int, string, Severity, bool> add)
        {
            var key = doc.Paragraphs.Select(x => TextUtil.Key(
                (string.IsNullOrEmpty(x.ListPrefix) ? "" : x.ListPrefix.Trim() + " ") + x.Text)).ToList();
            int Find(string rx, int from = 0, int to = int.MaxValue)
            {
                for (int i = from; i < key.Count && i < to; i++)
                    if (System.Text.RegularExpressions.Regex.IsMatch(key[i], rx)) return i;
                return -1;
            }
            void Need(string rx, string name, Severity sev = Severity.Warning)
            {
                if (Find(rx) < 0) add(StructureGroup, -1, $"Thiếu \"{name}\" theo khung kế hoạch bài dạy (CV 5512 Phụ lục IV).", sev, false);
            }
            Need(@"^TEN BAI DAY", "TÊN BÀI DẠY");
            Need(@"^MON HOC", "Môn học/Hoạt động giáo dục; lớp");
            Need(@"^THOI GIAN THUC HIEN", "Thời gian thực hiện");
            Need(@"^I\s*[\.\)]\s*MUC TIEU", "I. Mục tiêu", Severity.Error);
            Need(@"^\d\s*[\.\)]\s*VE KIEN THUC", "1. Về kiến thức");
            Need(@"^\d\s*[\.\)]\s*VE NANG LUC", "2. Về năng lực");
            Need(@"^\d\s*[\.\)]\s*VE PHAM CHAT", "3. Về phẩm chất");
            Need(@"^II\s*[\.\)]\s*THIET BI DAY HOC", "II. Thiết bị dạy học và học liệu", Severity.Error);
            Need(@"^III\s*[\.\)]\s*TIEN TRINH DAY HOC", "III. Tiến trình dạy học", Severity.Error);

            var names = new[] { "Mở đầu", "Hình thành kiến thức mới", "Luyện tập", "Vận dụng" };
            var acts = Enumerable.Range(1, 4).Select(a => Find($@"^(\d\s*[\.\)]\s*)?HOAT DONG {a}\b")).ToArray();
            for (int a = 0; a < 4; a++)
            {
                if (acts[a] < 0)
                {
                    // Hoạt động vận dụng chỉ thực hiện với bài có nội dung phù hợp (Ghi chú 1)
                    add(StructureGroup, -1, $"Thiếu \"Hoạt động {a + 1}: {names[a]}\" trong Tiến trình dạy học (CV 5512 Phụ lục IV).",
                        a == 3 ? Severity.Info : Severity.Warning, false);
                    continue;
                }
                int end = acts.Skip(a + 1).Where(x => x > acts[a]).DefaultIfEmpty(key.Count).Min();
                var missing = Enumerable.Range(0, 4)
                    .Where(st => Find($@"^({(char)('A' + st)}\s*[\.\)]\s*)?{Steps[st]}\b", acts[a] + 1, end) < 0)
                    .Select(st => $"{(char)('a' + st)}) {new[] { "Mục tiêu", "Nội dung", "Sản phẩm", "Tổ chức thực hiện" }[st]}").ToList();
                if (missing.Count > 0)
                    add(StructureGroup, acts[a], $"Hoạt động {a + 1} thiếu: {string.Join(", ", missing)} (CV 5512 Phụ lục IV).", Severity.Warning, false);
            }
        }

        static void CheckPage(PageInfo pg, FormatProfile p, Action<int, int, string, Severity, bool> add, bool landscapeAllowed)
        {
            bool Is(float w, float h) => Math.Abs(pg.WidthMm - w) <= 2 && Math.Abs(pg.HeightMm - h) <= 2;
            if (Is(p.PageHeightMm, p.PageWidthMm))
            {
                if (!landscapeAllowed)
                    add(PaperGroup, -1, "Trang ngang — văn bản không có bảng phải trình bày theo chiều dọc khổ A4 (210×297 mm).", Severity.Error, true);
            }
            else if (!Is(p.PageWidthMm, p.PageHeightMm))
                add(PaperGroup, -1, $"Khổ giấy {pg.WidthMm:0}×{pg.HeightMm:0} mm — phải là A4 (210×297 mm).", Severity.Error, true);
            void M(string name, float v, float min, float max)
            {
                if (v < min - 0.5 || v > max + 0.5)
                    add(MarginGroup, -1, $"Lề {name} {v:0.#} mm — quy định {min:0}–{max:0} mm.", Severity.Warning, true);
            }
            M("trên", pg.TopMm, p.MinTopMm, p.MaxTopMm);
            M("dưới", pg.BottomMm, p.MinBottomMm, p.MaxBottomMm);
            M("trái", pg.LeftMm, p.MinLeftMm, p.MaxLeftMm);
            M("phải", pg.RightMm, p.MinRightMm, p.MaxRightMm);
            if (p.PageNumberInHeader && pg.HasPageNumberInFooter && !pg.HasPageNumberInHeader)
                add(MiscGroup, -1, "Số trang phải đặt giữa phần đầu trang (header), không đặt ở chân trang.", Severity.Warning, true);
            if (p.HidePageNumberOnFirstPage && (pg.HasPageNumberInHeader || pg.HasPageNumberInFooter) && !pg.DifferentFirstPage)
                add(MiscGroup, -1, "Không hiển thị số trang ở trang thứ nhất.", Severity.Warning, true);
        }

        static string Trim(string s) => s.Length > 60 ? s.Substring(0, 57) + "..." : s;
        static string Range(ElementStyle s) => s.MinSize == s.MaxSize ? $"cỡ {s.MinSize:0.#}" : $"cỡ {s.MinSize:0.#}–{s.MaxSize:0.#}";
        static string Name(Align a) => a == Align.Left ? "trái" : a == Align.Right ? "phải" : a == Align.Center ? "giữa" : "đều hai bên";

        public static string Describe(ElementKind k)
        {
            switch (k)
            {
                case ElementKind.QuocHieu: return "Quốc hiệu";
                case ElementKind.TieuNgu: return "Tiêu ngữ";
                case ElementKind.TieuDeDang: return "Tiêu đề Đảng";
                case ElementKind.TieuDeDoan: return "Tiêu đề Đoàn";
                case ElementKind.CoQuanChuQuan: return "Tên cơ quan chủ quản";
                case ElementKind.CoQuanBanHanh: return "Tên cơ quan ban hành";
                case ElementKind.DauSao: return "Dấu phân cách";
                case ElementKind.SoKyHieu: return "Số, ký hiệu";
                case ElementKind.TrichYeuCongVan: return "Trích yếu công văn";
                case ElementKind.DiaDanhNgayThang: return "Địa danh, ngày tháng";
                case ElementKind.TenLoai: return "Tên loại văn bản";
                case ElementKind.TrichYeu: return "Trích yếu";
                case ElementKind.KinhGui: return "Kính gửi";
                case ElementKind.CanCu: return "Căn cứ ban hành";
                case ElementKind.Phan: return "Phần";
                case ElementKind.Chuong: return "Chương";
                case ElementKind.Muc: return "Mục";
                case ElementKind.TenPhanChuong: return "Tên phần/chương/mục";
                case ElementKind.Dieu: return "Điều";
                case ElementKind.MucLaMa: return "Đề mục La Mã";
                case ElementKind.MucSo: return "Đề mục số";
                case ElementKind.MucChuCai: return "Đề mục chữ cái";
                case ElementKind.TableText: return "Nội dung bảng";
                case ElementKind.ChucVuKy: return "Quyền hạn, chức vụ người ký";
                case ElementKind.HoTenKy: return "Họ tên người ký";
                case ElementKind.NoiNhanTitle: return "\"Nơi nhận\"";
                case ElementKind.NoiNhanItem: return "Danh sách nơi nhận";
                case ElementKind.ThongTinDau: return "Thông tin đầu văn bản";
                case ElementKind.GhiChuMau: return "Ghi chú biểu mẫu";
                case ElementKind.GhiChuKy: return "Ghi chú chữ ký";
                case ElementKind.XacThuc: return "\"XÁC THỰC VĂN BẢN HỢP NHẤT\"";
                case ElementKind.CoQuanXacThuc: return "Tên cơ quan ký xác thực";
                default: return "Nội dung";
            }
        }
    }
}
