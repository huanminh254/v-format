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
    }

    /// <summary>Định dạng cần áp cho một đoạn. null = giữ nguyên.</summary>
    public class ParagraphPlan
    {
        public float? Size;
        public bool? Bold, Italic;
        public int BoldPrefixLength;      // >0: chỉ in đậm n ký tự đầu (Điều dài)
        public int UnderlinePrefixLength; // >0: gạch chân n ký tự đầu ("Nơi nhận")
        public bool ToUpper;
        public Align? Align;
        public float? FirstIndentPt;      // null = giữ (danh sách đánh số)
        public bool ResetSideIndents;
        public float SpaceBeforePt, SpaceAfterPt;
        public float? LineMultiple;       // null = dòng đơn
        public bool KeepWithNext;
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
        static readonly HashSet<ElementKind> SignRight = new HashSet<ElementKind> { ElementKind.ChucVuKy, ElementKind.HoTenKy };

        static string Raw(ParagraphInfo p) => p.Text.TrimEnd('\r', '\a', '\n');

        /// <summary>Khối đầu đang là các đoạn rời → bảng 2 cột. null nếu không cần / không an toàn.</summary>
        public static TablePlan PlanHeader(ClassifiedDocument cd, FormatProfile p)
        {
            var k = cd.Kinds; var ps = cd.Paragraphs;
            int anchor = -1;
            for (int i = 0; i < k.Count && i < 20 && anchor < 0; i++)
                if (k[i] == ElementKind.QuocHieu || k[i] == ElementKind.TieuDeDang || k[i] == ElementKind.TieuDeDoan) anchor = i;
            if (anchor < 0 || ps[anchor].InTable) return null;

            int first = -1, last = -1;
            for (int i = 0; i < k.Count && i < 25; i++)
            {
                if (HeaderLeft.Contains(k[i]) || HeaderRight.Contains(k[i]))
                {
                    if (ps[i].InTable) return null;
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
                (HeaderLeft.Contains(k[i]) ? plan.Left : plan.Right).Add(Raw(ps[i]));
            }
            // Đảng / Đoàn: dấu sao ngay dưới tên cơ quan
            if (p.UseDauSao && !plan.Left.Any(x => TextUtil.Clean(x).Trim('*').Length == 0 && x.Contains("*")))
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
            var idx = Enumerable.Range(0, k.Count).Where(i => SignLeft.Contains(k[i]) || SignRight.Contains(k[i])).ToList();
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
                else if (k[i] == ElementKind.ChucVuKy) plan.Right.Add(raw);
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

        public static ParagraphPlan PlanParagraph(ElementKind kind, ParagraphInfo info, FormatProfile p, bool inList = false)
        {
            var st = p.StyleFor(kind);
            var text = TextUtil.Clean(info.Text);
            var plan = new ParagraphPlan();
            if (kind == ElementKind.Empty)
            {
                plan.Size = info.InTable ? st.Size : p.BodySize;
                return plan; // dòng đơn, khoảng cách 0
            }
            if (kind == ElementKind.TableText)
            {
                if (info.FontSize > 0 && (info.FontSize > st.MaxSize || info.FontSize < st.MinSize)) plan.Size = st.Size;
                plan.FirstIndentPt = 0;
                return plan;
            }

            plan.Size = st.Size;
            if (kind != ElementKind.Body)
            {
                if (kind == ElementKind.Dieu && text.Length > 90)
                {
                    plan.Bold = false;
                    var m = RxDieuPrefix.Match(info.Text);
                    if (m.Success) plan.BoldPrefixLength = m.Length;
                }
                else plan.Bold = st.Bold || kind == ElementKind.TrichYeu && TextUtil.IsAllUpper(text);
                plan.Italic = st.Italic;
            }
            if (kind == ElementKind.NoiNhanTitle && p.NoiNhanUnderline) plan.UnderlinePrefixLength = Math.Min(8, text.Length);
            plan.ToUpper = st.UpperCase && !TextUtil.IsAllUpper(text);
            plan.Align = st.Align;
            if (!inList)
            {
                plan.ResetSideIndents = true;
                plan.FirstIndentPt = info.InTable ? 0 : st.FirstIndentCm * 28.3465f;
            }
            plan.SpaceBeforePt = info.InTable && kind == ElementKind.NoiNhanTitle ? 0 : st.SpaceBeforePt;
            plan.SpaceAfterPt = info.InTable ? Math.Min(st.SpaceAfterPt, 3) : st.SpaceAfterPt;
            plan.LineMultiple = st.KeepLineSpacingSingle || info.InTable ? (float?)null : p.LineSpacing;
            plan.KeepWithNext = kind == ElementKind.TenLoai || kind == ElementKind.Chuong || kind == ElementKind.Phan
                                || kind == ElementKind.TenPhanChuong || kind == ElementKind.MucLaMa || kind == ElementKind.MucSo;
            return plan;
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
                        sign.Add(info.TableIndex); break;
                }
            }
            return (header, sign);
        }
    }
}
