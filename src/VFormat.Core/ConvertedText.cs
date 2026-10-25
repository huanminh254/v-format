using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

namespace VFormat.Core
{
    public enum LineBreakAction { Join, Split, Remove }

    /// <summary>
    /// Dọn văn bản chuyển từ PDF (hoặc gõ bằng Shift+Enter): mỗi dòng là một ngắt dòng cứng, khối đầu nhồi trong một đoạn.
    /// Dòng kết thúc bằng ngắt dòng cứng trong đoạn căn đều hai bên bị Word giãn chữ hết dòng ("BÁO ....... CÁO").
    /// </summary>
    public static class ConvertedText
    {
        public const char LineBreak = '\v';

        /// <summary>Dòng ngắn hơn ngưỡng này kết thúc bằng ngắt dòng là ngắt có chủ ý (dòng PDF bị cắt thường gần hết khổ).</summary>
        public const int ShortLine = 45;

        static readonly Regex RxBullet = new Regex(@"^[\-\+•–]\s");
        static readonly Regex RxListStart = new Regex(@"^((\d{1,2}|[IVXLC]{1,5}|[a-zđ])[\.\)]\s|\d{1,2}\s+\p{Lu})");
        static readonly Regex RxDecorLine = new Regex(@"^[_\-=\.…\*]{3,}$");
        static readonly Regex RxSentenceEnd = new Regex(@"[\.!?…][\)""”’]?$|[:;]$");

        /// <summary>Quyết định cho một ngắt dòng cứng giữa hai dòng: nối bằng khoảng trắng, tách thành đoạn mới, hoặc bỏ.</summary>
        public static LineBreakAction Decide(string before, string after)
        {
            var a = TextUtil.Clean(before);
            var b = TextUtil.Clean(after);
            if (a.Length == 0 || b.Length == 0) return LineBreakAction.Remove;
            if (RxBullet.IsMatch(b) || RxListStart.IsMatch(b) || RxDecorLine.IsMatch(b)) return LineBreakAction.Split;
            if (a.Length < ShortLine) return LineBreakAction.Split;
            if (RxSentenceEnd.IsMatch(a) && (char.IsUpper(b[0]) || b[0] == '(')) return LineBreakAction.Split;
            return LineBreakAction.Join;
        }

        /// <summary>Quyết định cho từng ngắt dòng cứng trong đoạn, theo thứ tự xuất hiện.</summary>
        public static LineBreakAction[] Plan(string raw)
        {
            var lines = (raw ?? "").TrimEnd('\r', '\a').Split(LineBreak);
            bool Blank(int i) => TextUtil.Clean(lines[i]).Length == 0;
            var actions = new LineBreakAction[lines.Length - 1];
            for (int i = 0; i < actions.Length; i++)
            {
                if (!Blank(i) && !Blank(i + 1)) { actions[i] = Decide(lines[i], lines[i + 1]); continue; }
                // ngắt dòng liền nhau (dòng trống ở giữa): ranh giới đoạn — tách một lần ở ngắt đầu tiên, bỏ các ngắt còn lại
                bool textBefore = Enumerable.Range(0, i + 1).Any(j => !Blank(j));
                bool textAfter = Enumerable.Range(i + 1, lines.Length - i - 1).Any(j => !Blank(j));
                bool firstOfRun = !Blank(i);
                actions[i] = textBefore && textAfter && firstOfRun ? LineBreakAction.Split : LineBreakAction.Remove;
            }
            return actions;
        }

        /// <summary>Số ký tự khoảng trắng / tab liền sau vị trí <paramref name="from"/> (để bỏ ở đầu đoạn mới tách).</summary>
        public static int LeadingBlank(string s, int from)
        {
            int n = 0;
            while (from + n < s.Length && (s[from + n] == ' ' || s[from + n] == '\t')) n++;
            return n;
        }

        /// <summary>Văn bản sau khi áp kế hoạch: các đoạn mới (đã nối dòng), giữ tab và ký tự gốc trong từng dòng.</summary>
        public static List<string> Reflow(string raw)
        {
            var lines = (raw ?? "").TrimEnd('\r', '\a').Split(LineBreak);
            var plan = Plan(raw);
            var result = new List<string> { lines[0] };
            for (int i = 0; i < plan.Length; i++)
            {
                if (plan[i] == LineBreakAction.Split) result.Add(lines[i + 1].TrimStart(' ', '\t'));
                else
                {
                    var last = result[result.Count - 1];
                    var sep = plan[i] == LineBreakAction.Join && !last.EndsWith(" ") && !lines[i + 1].StartsWith(" ") ? " " : "";
                    result[result.Count - 1] = last + sep + lines[i + 1];
                }
            }
            for (int i = 1; i < result.Count; i++) result[i] = result[i].TrimStart(' ', '\t');
            return result;
        }

        static readonly Regex RxQuocHieu = new Regex(@"C[ỘO]NG\s+H[ÒO]À?A?\s+X[ÃA]\s+H[ỘO]I\s+CH[ỦU]\s+NGH[ĨI]A\s+VI[ỆE]T\s+NAM",
            RegexOptions.IgnoreCase);
        static readonly Regex RxTieuNgu = new Regex(@"Độc\s+lập\s*[\-–—]\s*Tự\s+do\s*[\-–—]\s*Hạnh\s+phúc", RegexOptions.IgnoreCase);
        static readonly Regex RxFragmentSep = new Regex(@"\t+| {2,}");
        static readonly Regex RxFootnoteMark = new Regex(@"^\(?\d\)?(?=\p{Lu})");

        /// <summary>
        /// Khối đầu nhồi trong một đoạn ("1CỤC CẢNH SÁT GIAO      CỘNG HÒA XÃ HỘI... ⇥⇥THÔNG⇥ Độc lập - Tự do - Hạnh phúc ↵⇥2PHÒNG 6")
        /// → các dòng "tên cơ quan ⇥ Quốc hiệu", "tên cơ quan ⇥ Tiêu ngữ" (dạng khối đầu gõ tab mà Classifier nhận được).
        /// Trả về null nếu đoạn không phải khối đầu bị nhồi.
        /// </summary>
        public static string[] SplitMergedHeader(string raw)
        {
            raw = (raw ?? "").TrimEnd('\r', '\a').Normalize(System.Text.NormalizationForm.FormC);
            var qh = RxQuocHieu.Match(raw);
            if (!qh.Success) return null;
            var tn = RxTieuNgu.Match(raw);
            var right = new List<(int pos, string text)> { (qh.Index, TextNormalizer.QuocHieuText) };
            if (tn.Success) right.Add((tn.Index, TextNormalizer.TieuNguText));

            var rest = raw;
            foreach (var m in new[] { tn, qh }.Where(x => x.Success).OrderByDescending(x => x.Index))
                rest = rest.Substring(0, m.Index) + "\t" + rest.Substring(m.Index + m.Length);

            var left = new List<string>();
            foreach (var line in rest.Split(LineBreak))
            {
                var frags = RxFragmentSep.Split(line).Select(f => RxFootnoteMark.Replace(TextUtil.Clean(f), ""))
                    .Where(f => f.Length > 0).ToList();
                if (frags.Count == 0) continue;
                if (!frags.All(TextUtil.IsAllUpper)) return null;
                left.Add(string.Join(" ", frags));
            }
            if (left.Count == 0) return null;

            var rights = right.OrderBy(x => x.pos).Select(x => x.text).ToList();
            var lines = new List<string>();
            for (int i = 0; i < System.Math.Max(left.Count, rights.Count); i++)
            {
                var l = i < left.Count ? left[i] : "";
                var r = i < rights.Count ? rights[i] : "";
                lines.Add(l.Length > 0 && r.Length > 0 ? l + "\t" + r : l + r);
            }
            return lines.ToArray();
        }

        /// <summary>Số trang của bản PDF sót lại thành một đoạn riêng ("2") ở đầu trang.</summary>
        public static bool IsPageNumberArtifact(string text) => Regex.IsMatch(TextUtil.Clean(text), @"^\d{1,3}$");

        /// <summary>Văn bản có dấu hiệu chuyển từ PDF: nhiều đoạn ngoài bảng chứa ngắt dòng cứng.</summary>
        public static bool LooksConverted(IEnumerable<ParagraphInfo> paragraphs) =>
            paragraphs.Count(p => !p.InTable && p.Text.IndexOf(LineBreak) >= 0 && TextUtil.Clean(p.Text).Length > 0) >= 3;
    }
}
