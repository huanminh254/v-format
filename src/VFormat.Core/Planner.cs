using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

namespace VFormat.Core
{
    /// <summary>Kế hoạch dàn một khối (khối đầu / chữ ký) thành bảng 2 cột thay cho các đoạn [First..Last].</summary>
    public class TablePlan
    {
        public int First, Last;
        public List<string> Left = new List<string>();
        public List<string> Right = new List<string>();
        public float LeftMm, RightMm;
        public List<int> Remove = new List<int>();   // đoạn khác (ngoài First..Last) đã đưa vào bảng, cần xóa
    }

    /// <summary>Định dạng cần áp cho một đoạn. null = giữ nguyên.</summary>
    public class ParagraphPlan
    {
        public float? Size;
        public bool? Bold, Italic;
        public int BoldPrefixLength;      // >0: chỉ in đậm n ký tự đầu (Điều dài)
        public int UnderlinePrefixLength; // >0: gạch chân n ký tự đầu ("Nơi nhận")
        public int ItalicPrefixLength;    // >0: in nghiêng n ký tự đầu ("Kính gửi:" của Đảng, Đoàn)
        public bool ToUpper;
        public Align? Align;
        public float? FirstIndentPt;      // null = giữ (danh sách đánh số)
        public bool ResetSideIndents;
        public float? LeftIndentPt;       // đặt thụt lề trái này, thụt lề phải về 0 (khối ký ngoài bảng ở nửa phải trang)
        public float SpaceBeforePt, SpaceAfterPt;
        public float? LineMultiple;       // null = dòng đơn (hoặc Exactly nếu LineExactPt có giá trị)
        public float? LineExactPt;        // giãn dòng Exactly (Đảng)
        public bool KeepWithNext;
        public bool ResetColor;           // đưa màu chữ về tự động (đen)
        public bool ClearUnderline;       // bỏ gạch chân (chữ trong bảng)
        public bool BorderBottom;         // đường kẻ ngang hết chiều rộng dưới đoạn (ngăn phần ký xác thực, NĐ 250)
    }

    /// <summary>
    /// Các quyết định dàn trang và định dạng — dùng chung cho add-in Word và công cụ dòng lệnh,
    /// để hai nơi cho ra cùng một kết quả.
    /// </summary>
    public static class Planner
    {
        static readonly HashSet<ElementKind> HeaderLeft = new HashSet<ElementKind>
            { ElementKind.CoQuanChuQuan, ElementKind.CoQuanBanHanh, ElementKind.DauSao, ElementKind.SoKyHieu, ElementKind.TrichYeuCongVan };
        static readonly HashSet<ElementKind> HeaderRight = new HashSet<ElementKind>
            { ElementKind.QuocHieu, ElementKind.TieuNgu, ElementKind.TieuDeDang, ElementKind.TieuDeDoan, ElementKind.DiaDanhNgayThang };
        static readonly HashSet<ElementKind> SignLeft = new HashSet<ElementKind> { ElementKind.NoiNhanTitle, ElementKind.NoiNhanItem };
        static readonly HashSet<ElementKind> SignRight = new HashSet<ElementKind>
            { ElementKind.ChucVuKy, ElementKind.HoTenKy, ElementKind.GhiChuKy, ElementKind.XacThuc, ElementKind.CoQuanXacThuc };

        /// <summary>Đoạn thuộc cột phải khối chữ ký (kể cả ngày tháng ngay dưới "XÁC THỰC VĂN BẢN HỢP NHẤT").</summary>
        static bool IsSignRight(ClassifiedDocument cd, int i)
        {
            var k = cd.Kinds;
            if (SignRight.Contains(k[i])) return true;
            if (k[i] != ElementKind.DiaDanhNgayThang) return false;
            int j = i - 1;
            while (j >= 0 && k[j] == ElementKind.Empty) j--;
            if (j >= 0 && k[j] == ElementKind.XacThuc) return true;
            // ngày tháng đặt ngay trên chức vụ người ký (báo cáo, đơn, biểu mẫu)
            j = i + 1;
            while (j < k.Count && k[j] == ElementKind.Empty) j++;
            return j < k.Count && k[j] == ElementKind.ChucVuKy && !cd.Paragraphs[i].InTable;
        }

        static bool IsConsolidatedLabel(string text) => TextUtil.Key(text) == "VAN BAN HOP NHAT";

        static string Raw(ParagraphInfo p) => p.Text.TrimEnd('\r', '\a', '\n');

        /// <summary>Khối đầu đang là các đoạn rời → bảng 2 cột. null nếu không cần / không an toàn.</summary>
        public static TablePlan PlanHeader(ClassifiedDocument cd, FormatProfile p)
        {
            var k = cd.Kinds; var ps = cd.Paragraphs;
            int anchor = -1;
            for (int i = 0; i < k.Count && i < 20 && anchor < 0; i++)
                if (k[i] == ElementKind.QuocHieu || k[i] == ElementKind.TieuDeDang || k[i] == ElementKind.TieuDeDoan) anchor = i;
            if (anchor < 0 || ps[anchor].InTable) return null;
            bool hasSeparator = false;

            int first = -1, last = -1;
            for (int i = 0; i < k.Count && i < 25; i++)
            {
                if (HeaderLeft.Contains(k[i]) || HeaderRight.Contains(k[i]))
                {
                    if (ps[i].InTable) return null;
                    if (k[i] == ElementKind.DauSao) hasSeparator = true;
                    if (first < 0) first = i;
                    last = i;
                }
                else if (k[i] != ElementKind.Empty && first >= 0) break;
            }
            if (first < 0) return null;

            var plan = new TablePlan { First = first, Last = last, LeftMm = p.HeaderLeftColMm, RightMm = p.HeaderRightColMm };
            for (int i = first; i <= last; i++)
            {
                if (k[i] == ElementKind.Empty) continue;
                if (cd.Splits.TryGetValue(i, out var sp))
                {
                    // dòng gõ tab hai cột: tách về đúng ô
                    if (sp.LeftKind != ElementKind.Empty) plan.Left.Add(sp.Left);
                    if (sp.LeftKind == ElementKind.DauSao) hasSeparator = true;
                    plan.Right.Add(sp.Right);
                    continue;
                }
                (HeaderLeft.Contains(k[i]) ? plan.Left : plan.Right).Add(Raw(ps[i]));
            }
            // Văn bản hợp nhất: "VĂN BẢN HỢP NHẤT" đặt ở ô trái, ngay trên số, ký hiệu (NĐ 250/2026 Phụ lục mục I.2)
            int label = ConsolidatedLabelOutsideHeader(cd, p);
            if (label > last && !plan.Left.Any(IsConsolidatedLabel))
            {
                int so = plan.Left.FindIndex(x => TextUtil.Key(x).StartsWith("SO"));
                plan.Left.Insert(so < 0 ? plan.Left.Count : so, TextUtil.Clean(ps[label].Text));
                plan.Remove.Add(label);
            }
            // Đảng / Đoàn: dấu sao ngay dưới tên cơ quan
            if (p.UseDauSao && !hasSeparator)
            {
                int so = plan.Left.FindIndex(x => TextUtil.Key(x).StartsWith("SO"));
                plan.Left.Insert(so < 0 ? plan.Left.Count : so, p.DauSaoText);
            }
            return plan;
        }

        /// <summary>Khối chữ ký + nơi nhận đang là các đoạn rời → bảng 2 cột (trái: nơi nhận, phải: chữ ký).</summary>
        public static TablePlan PlanSign(ClassifiedDocument cd, FormatProfile p)
        {
            var k = cd.Kinds; var ps = cd.Paragraphs;
            var idx = Enumerable.Range(0, k.Count).Where(i => SignLeft.Contains(k[i]) || IsSignRight(cd, i)).ToList();
            if (!idx.Any(i => k[i] == ElementKind.ChucVuKy)) return null;
            if (idx.Any(i => ps[i].InTable)) return null;
            int first = idx.First(), last = idx.Last();
            if (last - first > 40) return null;

            var plan = new TablePlan { First = first, Last = last, LeftMm = p.SignLeftColMm, RightMm = p.SignRightColMm };
            bool nameSeen = false;
            for (int i = first; i <= last; i++)
            {
                var raw = Raw(ps[i]);
                var t = TextUtil.Clean(raw);
                if (k[i] == ElementKind.Empty) continue;
                if (SignLeft.Contains(k[i])) plan.Left.Add(raw);
                else if (k[i] == ElementKind.ChucVuKy || k[i] == ElementKind.XacThuc || k[i] == ElementKind.CoQuanXacThuc
                         || k[i] == ElementKind.DiaDanhNgayThang) plan.Right.Add(raw);
                else if (k[i] == ElementKind.HoTenKy)
                {
                    for (int b = 0; b < 4; b++) plan.Right.Add(""); // chừa chỗ ký, đóng dấu
                    plan.Right.Add(raw); nameSeen = true;
                }
                else if (t.StartsWith("(") && t.EndsWith(")")) plan.Right.Add(raw);   // (Đã ký) / (Ký, ghi rõ họ tên)
                else return null; // có nội dung khác xen giữa: không tự dàn
            }
            if (!nameSeen) for (int b = 0; b < 4; b++) plan.Right.Add("");
            if (plan.Left.Count == 0) plan.Left.Add("");
            return plan;
        }

        /// <summary>Đoạn được sửa lại chữ (Quốc hiệu, số ký hiệu, nơi nhận...; căn cứ của Đảng).</summary>
        public static bool NeedsTextFix(ElementKind k, FormatProfile p) =>
            IsFormal(k) || k == ElementKind.GhiChuMau || k == ElementKind.CanCu && p.CanCuDash || k == ElementKind.Phan || k == ElementKind.Chuong
            || k == ElementKind.ThongTinDau || k == ElementKind.TrichYeu && p.Kind == StandardKind.GiaoDuc_CV5512;

        /// <summary>Công văn không ghi tên loại: dòng "công văn" đứng riêng (văn bản có trích yếu "V/v") cần bỏ; -1 nếu không có.</summary>
        public static int StrayCongVanLabel(ClassifiedDocument cd)
        {
            int v = cd.Kinds.IndexOf(ElementKind.TrichYeuCongVan);
            if (v < 0) return -1;
            for (int i = 0; i < cd.Kinds.Count && i <= v + 2; i++)
                if (!cd.Paragraphs[i].InTable && TextUtil.Key(cd.Paragraphs[i].Text) == "CONG VAN") return i;
            return -1;
        }

        /// <summary>
        /// Văn bản hợp nhất (NĐ 250/2026): dòng "VĂN BẢN HỢP NHẤT" đang nằm giữa trang (trên tên loại) thay vì ở ô trái khối đầu;
        /// -1 nếu không có.
        /// </summary>
        public static int ConsolidatedLabelOutsideHeader(ClassifiedDocument cd, FormatProfile p)
        {
            if (p.Kind != StandardKind.PhapDien_ND250) return -1;
            for (int i = 0; i < cd.Kinds.Count && i < 40; i++)
                if (cd.Kinds[i] == ElementKind.TenLoai && !cd.Paragraphs[i].InTable && IsConsolidatedLabel(cd.Paragraphs[i].Text)) return i;
            return -1;
        }

        /// <summary>
        /// Khối đầu đã là bảng: chuyển dòng "VĂN BẢN HỢP NHẤT" vào ô trái, ngay trên số, ký hiệu.
        /// Trả về (đoạn "VĂN BẢN HỢP NHẤT", đoạn số ký hiệu) hoặc (-1, -1).
        /// </summary>
        public static (int from, int before) PlanConsolidatedLabelMove(ClassifiedDocument cd, FormatProfile p)
        {
            int from = ConsolidatedLabelOutsideHeader(cd, p), so = cd.Kinds.IndexOf(ElementKind.SoKyHieu);
            if (from < 0 || so < 0 || so > from || !cd.Paragraphs[so].InTable) return (-1, -1);
            return (from, so);
        }

        /// <summary>
        /// NĐ 250/2026 mục XIII: phần ký xác thực ngăn với nội dung bằng đường kẻ ngang, dài, nét liền.
        /// Trả về đoạn nội dung cuối cùng (kẻ viền dưới) hoặc -1.
        /// </summary>
        public static int SignSeparatorIndex(ClassifiedDocument cd, FormatProfile p)
        {
            if (p.Kind != StandardKind.PhapDien_ND250) return -1;
            var k = cd.Kinds;
            int x = k.IndexOf(ElementKind.XacThuc);
            if (x < 0) return -1;
            int start = x;
            for (int i = x - 1; i >= 0; i--)
            {
                if (k[i] == ElementKind.Empty) continue;
                if (SignLeft.Contains(k[i]) || IsSignRight(cd, i)) { start = i; continue; }
                break;
            }
            // bảng chữ ký: đoạn đứng trước bảng
            if (cd.Paragraphs[start].InTable)
            {
                int table = cd.Paragraphs[start].TableIndex;
                while (start > 0 && cd.Paragraphs[start - 1].InTable && cd.Paragraphs[start - 1].TableIndex == table) start--;
            }
            for (int i = start - 1; i >= 0; i--)
            {
                if (k[i] == ElementKind.Empty) continue;
                return cd.Paragraphs[i].InTable ? -1 : i;
            }
            return -1;
        }

        /// <summary>
        /// Trích yếu công văn ("V/v ...") nằm ngoài khối đầu dạng bảng → chuyển vào ô trái, ngay dưới số, ký hiệu.
        /// Trả về (đoạn trích yếu, đoạn số ký hiệu) hoặc (-1, -1).
        /// </summary>
        public static (int from, int after) PlanCongVanSubjectMove(ClassifiedDocument cd)
        {
            int v = cd.Kinds.IndexOf(ElementKind.TrichYeuCongVan), s = cd.Kinds.IndexOf(ElementKind.SoKyHieu);
            if (v < 0 || s < 0 || v < s || cd.Paragraphs[v].InTable || !cd.Paragraphs[s].InTable) return (-1, -1);
            return (v, s);
        }

        /// <summary>
        /// Độ rộng cột (mm) của bảng nội dung theo độ dài chữ trong từng cột: cột chữ dài rộng hơn cột số, không cột nào dưới 12 mm.
        /// null nếu các hàng có số cột khác nhau (ô gộp).
        /// </summary>
        public static float[] PlanColumnWidths(ClassifiedDocument cd, int tableIndex, float availableMm)
        {
            var cells = cd.Paragraphs.Where(x => x.InTable && x.TableIndex == tableIndex).ToList();
            if (cells.Count == 0) return null;
            int cols = cells.Max(x => x.CellColumn);
            if (cols < 2 || cells.GroupBy(x => x.CellRow).Any(r => r.Max(x => x.CellColumn) != cols)) return null;
            var weight = new float[cols];
            foreach (var g in cells.GroupBy(x => (x.CellRow, x.CellColumn)))
            {
                int len = TextUtil.Clean(string.Join(" ", g.Select(x => x.Text))).Length;
                int c = g.Key.CellColumn - 1;
                weight[c] = Math.Max(weight[c], Math.Min(Math.Max(len, 4), 60));
            }
            const float min = 12;
            var w = weight.Select(x => availableMm * x / weight.Sum()).ToArray();
            for (int pass = 0; pass < 3; pass++)
            {
                float deficit = w.Where(x => x < min).Sum(x => min - x), spare = w.Where(x => x > min).Sum(x => x - min);
                if (deficit <= 0 || spare <= 0) break;
                w = w.Select(x => x < min ? min : x - (x - min) / spare * deficit).ToArray();
            }
            return w;
        }

        /// <summary>Như trên, thêm dòng dấu phân cách dưới tên cơ quan của Đảng ("*"), Đoàn ("***").</summary>
        public static bool NeedsTextFix(ClassifiedDocument cd, int i, FormatProfile p) =>
            NeedsTextFix(cd.Kinds[i], p) || p.UseDauSao && i == HeaderSeparatorIndex(cd);

        /// <summary>Chữ đúng của đoạn i (dùng chung cho chuẩn hóa và rà soát).</summary>
        public static string FixText(ClassifiedDocument cd, int i, FormatProfile p, string clean)
        {
            if (p.UseDauSao && i == HeaderSeparatorIndex(cd)) return p.DauSaoText;
            var kind = cd.Kinds[i];
            return TextNormalizer.Fix(kind, clean, p, i == cd.Kinds.LastIndexOf(kind));
        }

        /// <summary>Dòng phân cách ngay dưới tên cơ quan ban hành (Đảng "*", Đoàn "***"); -1 nếu không có.</summary>
        public static int HeaderSeparatorIndex(ClassifiedDocument cd)
        {
            var k = cd.Kinds;
            for (int i = 1; i < k.Count && i < 25; i++)
            {
                if (k[i] != ElementKind.DauSao) continue;
                int prev = i - 1;
                while (prev >= 0 && k[prev] == ElementKind.Empty) prev--;
                if (prev >= 0 && (k[prev] == ElementKind.CoQuanBanHanh || k[prev] == ElementKind.CoQuanChuQuan)) return i;
            }
            return -1;
        }

        /// <summary>
        /// Có bảng nội dung (không tính bảng khối đầu, khối chữ ký). Văn bản có bảng mới được trình bày trang ngang
        /// (HD 05 mục IV.4.3, HD 29 mục II.2); văn bản khác đưa về A4 dọc.
        /// </summary>
        public static bool HasContentTable(ClassifiedDocument cd)
        {
            var (header, sign) = FormalTables(cd);
            return cd.Paragraphs.Any(x => x.InTable && !header.Contains(x.TableIndex) && !sign.Contains(x.TableIndex));
        }

        /// <summary>
        /// In đậm cả đoạn hay không. Đảng, Đoàn: dòng quyền hạn ký ("T/M BAN THƯỜNG VỤ") đậm, dòng chức vụ ("PHÓ BÍ THƯ") đứng.
        /// </summary>
        public static bool WantBold(ElementKind kind, string text, FormatProfile p)
        {
            if (kind == ElementKind.ChucVuKy && !p.SignerTitleBold && !Classifier.HasSignAuthority(text)) return false;
            return p.StyleFor(kind).Bold || kind == ElementKind.TrichYeu && TextUtil.IsAllUpper(text);
        }

        static readonly Regex RxKinhGuiPrefix = new Regex(@"^\s*Kính\s+gửi\s*:");

        /// <summary>
        /// Đảng, Đoàn: dưới tên loại và trích yếu có dòng gạch nối ("-----"). Trả về chỉ số đoạn cần chèn dòng gạch ngay sau,
        /// -1 nếu quy chuẩn không yêu cầu, không có tên loại hoặc đã có.
        /// </summary>
        public static int PlanTitleDashes(ClassifiedDocument cd, FormatProfile p)
        {
            if (string.IsNullOrEmpty(p.TitleDashes)) return -1;
            var k = cd.Kinds;
            int anchor = k.LastIndexOf(ElementKind.TrichYeu);
            if (anchor < 0) anchor = k.IndexOf(ElementKind.TenLoai);
            if (anchor < 0 || cd.Paragraphs[anchor].InTable) return -1;
            int next = anchor + 1;
            while (next < k.Count && k[next] == ElementKind.Empty) next++;
            return next < k.Count && k[next] == ElementKind.DauSao ? -1 : anchor;
        }

        public static bool IsFormal(ElementKind k)
        {
            switch (k)
            {
                case ElementKind.QuocHieu: case ElementKind.TieuNgu: case ElementKind.TieuDeDang: case ElementKind.TieuDeDoan:
                case ElementKind.SoKyHieu: case ElementKind.DiaDanhNgayThang: case ElementKind.TrichYeuCongVan:
                case ElementKind.KinhGui: case ElementKind.NoiNhanTitle: case ElementKind.NoiNhanItem: case ElementKind.ChucVuKy:
                case ElementKind.HoTenKy: case ElementKind.CoQuanBanHanh: case ElementKind.CoQuanChuQuan: case ElementKind.Dieu:
                    return true;
            }
            return false;
        }

        public static bool IsBodyKind(ElementKind k) =>
            k == ElementKind.Body || k == ElementKind.CanCu || k == ElementKind.Dieu || k == ElementKind.MucLaMa ||
            k == ElementKind.MucSo || k == ElementKind.MucChuCai || k == ElementKind.TableText || k == ElementKind.TrichYeu;

        /// <summary>Đoạn nội dung đầu/cuối nằm ngoài bảng — phạm vi sửa khoảng trắng (không chạm khối đầu).</summary>
        public static (int first, int last) BodyRange(ClassifiedDocument cd)
        {
            int first = -1, last = -1;
            for (int i = 0; i < cd.Kinds.Count; i++)
                if (!cd.Paragraphs[i].InTable && IsBodyKind(cd.Kinds[i])) { if (first < 0) first = i; last = i; }
            return (first, last);
        }

        /// <summary>Các đoạn cần kẻ đường bên dưới, kèm tỉ lệ độ dài so với dòng chữ.</summary>
        public static List<(int index, float ratio)> PlanLines(ClassifiedDocument cd, FormatProfile p)
        {
            var list = new List<(int, float)>();
            var k = cd.Kinds;
            int lastTrichYeu = k.LastIndexOf(ElementKind.TrichYeu);
            for (int i = 0; i < k.Count; i++)
            {
                float ratio = 0;
                switch (k[i])
                {
                    case ElementKind.TieuNgu: case ElementKind.TieuDeDang: case ElementKind.TieuDeDoan:
                        ratio = p.HeaderTitleLineRatio; break;
                    case ElementKind.CoQuanBanHanh:
                        ratio = p.CoQuanLineRatio; break;
                    case ElementKind.CoQuanXacThuc:
                        ratio = 0.4f; break;   // NĐ 250/2026 mục XIII: kẻ dưới tên cơ quan ký xác thực
                    case ElementKind.TrichYeu:
                        if (i == lastTrichYeu) ratio = p.TrichYeuLineRatio; break;
                    case ElementKind.TenLoai:
                        // tên loại không có trích yếu (Giấy mời, Giấy giới thiệu...) thì kẻ dưới tên loại
                        if (lastTrichYeu < 0 && i == k.IndexOf(ElementKind.TenLoai)) ratio = p.TrichYeuLineRatio; break;
                }
                if (ratio <= 0) continue;
                // Bỏ qua nếu ngay dưới đã có dòng "-----" hoặc "***"
                int next = i + 1;
                while (next < k.Count && k[next] == ElementKind.Empty) next++;
                if (next < k.Count && k[next] == ElementKind.DauSao) continue;
                list.Add((i, ratio));
            }
            return list;
        }

        static readonly Regex RxDieuPrefix = new Regex(@"^\s*Điều\s+\d+[a-z]?\.");

        /// <summary>Định dạng đoạn i, tính cả ngữ cảnh văn bản (dòng khối đầu gõ tab, đường kẻ trên phần ký xác thực).</summary>
        public static ParagraphPlan PlanParagraph(ClassifiedDocument cd, int i, FormatProfile p, bool inList = false)
        {
            var info = cd.Paragraphs[i];
            if (cd.Splits.ContainsKey(i))
            {
                // khối đầu chưa dàn bảng: giữ tab, căn trái, không in đậm / in hoa cả dòng (phần trái là tên cơ quan)
                return new ParagraphPlan
                {
                    ResetColor = true, Size = p.StyleFor(cd.Kinds[i]).Size, Align = Align.Left, FirstIndentPt = 0,
                    ResetSideIndents = true,
                };
            }
            var plan = PlanParagraph(cd.Kinds[i], info, p, inList, signArea: IsSignRight(cd, i));
            if (i == SignSeparatorIndex(cd, p)) plan.BorderBottom = true;
            return plan;
        }

        public const float PtPerMm = 72f / 25.4f;

        /// <param name="signArea">đoạn thuộc cột phải khối chữ ký (kể cả ngày tháng ngay trên chức vụ người ký)</param>
        public static ParagraphPlan PlanParagraph(ElementKind kind, ParagraphInfo info, FormatProfile p, bool inList = false,
            bool signArea = false)
        {
            var st = p.StyleFor(kind);
            var text = TextUtil.Clean(info.Text);
            var plan = new ParagraphPlan { ResetColor = true };
            if (kind == ElementKind.Empty)
            {
                plan.Size = info.InTable ? st.Size : p.BodySize;
                return plan; // dòng đơn, khoảng cách 0
            }
            if (kind == ElementKind.GhiChuMau)
            {
                // ghi chú biểu mẫu: giữ nguyên cỡ chữ, vị trí; bị đẩy sang phải bằng thụt lề lớn (bản PDF) thì đặt sát lề phải
                if (!info.InTable && info.LeftIndentPt >= 100)
                {
                    plan.Align = Align.Right;
                    plan.ResetSideIndents = true;
                    plan.FirstIndentPt = 0;
                }
                return plan;
            }
            if (kind == ElementKind.TableText)
            {
                if (info.FontSize > 0 && (info.FontSize > st.MaxSize || info.FontSize < st.MinSize)) plan.Size = st.Size;
                plan.FirstIndentPt = 0;
                // chữ trong ô: bỏ đậm / nghiêng phủ cả ô và gạch chân (hàng tiêu đề được in đậm lại khi định dạng bảng)
                if (info.Bold == 1) plan.Bold = false;
                if (info.Italic == 1) plan.Italic = false;
                plan.ClearUnderline = true;
                return plan;
            }

            plan.Size = st.Size;
            if (kind == ElementKind.Body)
            {
                // lời văn kiểu chữ đứng: bỏ đậm / nghiêng phủ cả đoạn, giữ nhấn mạnh từng chữ
                if (info.Bold == 1) plan.Bold = false;
                if (info.Italic == 1) plan.Italic = false;
            }
            else
            {
                // Điều có tiêu đề ("Điều 5. Phạm vi điều chỉnh") in đậm cả dòng; điều viết liền nội dung
                // ("Điều 1. Ban hành kèm theo... ." — quyết định cá biệt) chỉ in đậm "Điều 1."
                if (kind == ElementKind.Dieu && IsDieuWithContent(text))
                {
                    plan.Bold = false;
                    var m = RxDieuPrefix.Match(info.Text);
                    if (m.Success) plan.BoldPrefixLength = m.Length;
                }
                else plan.Bold = WantBold(kind, text, p);
                plan.Italic = st.Italic;
                if (kind == ElementKind.KinhGui && p.KinhGuiPrefixItalic)
                {
                    var m = RxKinhGuiPrefix.Match(info.Text);
                    if (m.Success) plan.ItalicPrefixLength = m.Length;
                }
            }
            if (kind == ElementKind.NoiNhanTitle && p.NoiNhanUnderline) plan.UnderlinePrefixLength = Math.Min(8, text.Length);
            plan.ToUpper = st.UpperCase && !TextUtil.IsAllUpper(text);
            plan.Align = st.Align;
            // khối chữ ký không dàn được thành bảng: đặt vào nửa phải trang đúng vị trí cột phải của bảng chữ ký,
            // căn giữa trong cột (thụt lề trái / phải tùy tiện của bản gốc, bản PDF làm gãy dòng chức vụ)
            bool signRight = !info.InTable && (signArea || SignRight.Contains(kind) ||
                kind == ElementKind.DiaDanhNgayThang && info.LeftIndentPt >= 100);
            if (signRight)
            {
                plan.FirstIndentPt = 0;
                plan.LeftIndentPt = p.SignLeftColMm * PtPerMm;   // căn lề giữ theo quy chuẩn (ngày tháng Đảng, Đoàn: căn phải)
            }
            else if (!inList)
            {
                plan.ResetSideIndents = true;
                plan.FirstIndentPt = info.InTable ? 0 : st.FirstIndentCm * 28.3465f;
            }
            plan.SpaceBeforePt = info.InTable && kind == ElementKind.NoiNhanTitle ? 0 : st.SpaceBeforePt;
            plan.SpaceAfterPt = info.InTable ? Math.Min(st.SpaceAfterPt, 3) : st.SpaceAfterPt;
            bool single = st.KeepLineSpacingSingle || info.InTable;
            if (!single && p.LineExactPt > 0) plan.LineExactPt = p.LineExactPt;
            else plan.LineMultiple = single ? (float?)null : p.LineSpacing;
            plan.KeepWithNext = kind == ElementKind.TenLoai || kind == ElementKind.Chuong || kind == ElementKind.Phan
                                || kind == ElementKind.TenPhanChuong || kind == ElementKind.MucLaMa || kind == ElementKind.MucSo;
            return plan;
        }

        /// <summary>Phần sau "Điều N." là câu nội dung (kết thúc bằng dấu câu), không phải tiêu đề điều.</summary>
        public static bool IsDieuWithContent(string text)
        {
            var m = RxDieuPrefix.Match(text);
            if (!m.Success) return false;
            var rest = text.Substring(m.Length).Trim();
            return rest.Length > 0 && (Regex.IsMatch(rest, @"[\.;:]$|\./\.$") || rest.Length > 160);
        }

        /// <summary>Bảng nào là khối đầu (header) và khối chữ ký (sign).</summary>
        public static (HashSet<int> header, HashSet<int> sign) FormalTables(ClassifiedDocument cd)
        {
            var header = new HashSet<int>(); var sign = new HashSet<int>();
            for (int i = 0; i < cd.Kinds.Count; i++)
            {
                var info = cd.Paragraphs[i];
                if (!info.InTable) continue;
                switch (cd.Kinds[i])
                {
                    case ElementKind.QuocHieu: case ElementKind.TieuNgu: case ElementKind.TieuDeDang: case ElementKind.TieuDeDoan:
                    case ElementKind.CoQuanBanHanh: case ElementKind.SoKyHieu:
                        header.Add(info.TableIndex); break;
                    case ElementKind.ChucVuKy: case ElementKind.NoiNhanTitle: case ElementKind.HoTenKy:
                    case ElementKind.XacThuc: case ElementKind.CoQuanXacThuc:
                        sign.Add(info.TableIndex); break;
                }
            }
            return (header, sign);
        }
    }
}
