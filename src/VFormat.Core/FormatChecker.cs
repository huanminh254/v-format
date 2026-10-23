using System;
using System.Collections.Generic;
using System.Linq;

namespace VFormat.Core
{
    /// <summary>
    /// Rà soát thể thức. Mã lỗi gồm tiền tố quy chuẩn + nhóm 01..10:
    /// 01 khổ giấy · 02 lề trang · 03 phông chữ · 04 cỡ chữ · 05 kiểu chữ (đậm/nghiêng/in hoa)
    /// 06 căn lề/thụt dòng · 07 Quốc hiệu, tiêu ngữ, tiêu đề · 08 số ký hiệu, ngày tháng
    /// 09 chữ ký, nơi nhận · 10 thiếu thành phần, khoảng trắng, số trang.
    /// </summary>
    public static class FormatChecker
    {
        public const int PaperGroup = 1, MarginGroup = 2, FontGroup = 3, SizeGroup = 4, StyleGroup = 5,
            AlignGroup = 6, HeaderGroup = 7, NumberDateGroup = 8, SignGroup = 9, MiscGroup = 10;

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
            void Add(int group, int para, string msg, Severity sev = Severity.Warning, bool fix = true) =>
                issues.Add(new Issue { Code = Code(p, group), ParagraphIndex = para, Message = msg, Severity = sev, AutoFixable = fix });

            if (page != null) CheckPage(page, p, Add);

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
            }

            int lastNoiNhan = doc.Kinds.LastIndexOf(ElementKind.NoiNhanItem);
            for (int i = 0; i < doc.Paragraphs.Count; i++)
            {
                var kind = doc.Kinds[i];
                var para = doc.Paragraphs[i];
                if (kind == ElementKind.Empty) continue;
                var st = p.StyleFor(kind);
                var text = TextUtil.Clean(para.Text);

                if (para.FontName != null && !string.Equals(para.FontName, p.FontName, StringComparison.OrdinalIgnoreCase))
                    Add(FontGroup, i, $"Phông chữ \"{para.FontName}\" — phải dùng {p.FontName}.");

                if (kind != ElementKind.TableText && para.FontSize > 0 && (para.FontSize < st.MinSize - 0.01 || para.FontSize > st.MaxSize + 0.01))
                    Add(SizeGroup, i, $"{Describe(kind)}: cỡ chữ {para.FontSize:0.#} — quy định {Range(st)}.");

                if (kind != ElementKind.TableText && kind != ElementKind.Body)
                {
                    bool wantBold = st.Bold || kind == ElementKind.TrichYeu && TextUtil.IsAllUpper(text);
                    if (para.Bold >= 0 && (para.Bold == 1) != wantBold)
                        Add(StyleGroup, i, $"{Describe(kind)} phải {(wantBold ? "in đậm" : "không in đậm")}.");
                    if (para.Italic >= 0 && (para.Italic == 1) != st.Italic)
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

                if (!para.InTable && st.FirstIndentCm > 0 && Math.Abs(para.FirstIndentPt - st.FirstIndentCm * 28.3465f) > 3
                    && (kind == ElementKind.Body || kind == ElementKind.Dieu || kind == ElementKind.CanCu))
                    Add(AlignGroup, i, $"Thụt đầu dòng {para.FirstIndentPt / 28.3465f:0.##} cm — quy định {st.FirstIndentCm:0.##} cm.");

                // Nội dung chữ
                var fixedText = TextNormalizer.Fix(kind, text, p, i == lastNoiNhan);
                if (TextUtil.Clean(fixedText) != text)
                {
                    int group;
                    switch (kind)
                    {
                        case ElementKind.QuocHieu: case ElementKind.TieuNgu: case ElementKind.TieuDeDang: case ElementKind.TieuDeDoan:
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

        static void CheckPage(PageInfo pg, FormatProfile p, Action<int, int, string, Severity, bool> add)
        {
            if (Math.Abs(pg.WidthMm - p.PageWidthMm) > 2 || Math.Abs(pg.HeightMm - p.PageHeightMm) > 2)
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
                default: return "Nội dung";
            }
        }
    }
}
