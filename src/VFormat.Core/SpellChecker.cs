using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

namespace VFormat.Core
{
    public class SpellHit
    {
        public int ParagraphIndex;
        public int Start;          // vị trí trong chuỗi đoạn văn (chưa Clean)
        public int Length;
        public string Found;
        public string Suggestion;
        public string Reason;
        public override string ToString() => $"\"{Found}\" → \"{Suggestion}\" ({Reason})";
    }

    /// <summary>
    /// Kiểm tra chính tả offline: lỗi phụ âm/vần hay gặp, viết hoa tên cơ quan theo Phụ lục II NĐ 30,
    /// từ lặp và khoảng trắng. Bộ quy tắc có thể mở rộng bằng tệp người dùng (sai=đúng mỗi dòng).
    /// </summary>
    public class SpellChecker
    {
        class Rule
        {
            public Regex Pattern;
            public string Replacement;   // chuỗi thay thế Regex
            public string Reason;
        }

        readonly List<Rule> _rules = new List<Rule>();

        // Lỗi chính tả thường gặp: 403 cặp sinh từ typo_corrections_map.json
        // (nguồn: Từ điển tiếng Việt Viện Ngôn ngữ học, Nghị quyết UBTVQH).
        // Xem SpellCheckData.TypoPairs. Thêm cặp thủ công vào đây nếu cần ghi đè.
        static readonly (string wrong, string right)[] ExtraMistakes =
        {
            // Các cặp đã có trong TypoPairs — giữ lại chỗ này để dễ bổ sung thủ công khi phát hiện lỗi mới.
            ("qui định", "quy định"), ("qui chế", "quy chế"), ("qui trình", "quy trình"),
            ("qui mô", "quy mô"), ("qui hoạch", "quy hoạch"), ("quí", "quý"), ("lí do", "lý do"),
            ("lí luận", "lý luận"), ("kĩ thuật", "kỹ thuật"), ("kỉ luật", "kỷ luật"),
        };

        // Viết hoa tên cơ quan, chức danh (Phụ lục II NĐ 30/2020)
        static readonly (string pattern, string right)[] Capitalization =
        {
            (@"uỷ ban nhân dân", "Ủy ban nhân dân"), (@"ủy ban nhân dân", "Ủy ban nhân dân"), (@"Ủy Ban Nhân Dân", "Ủy ban nhân dân"),
            (@"Ủy Ban nhân dân", "Ủy ban nhân dân"), (@"Hội đồng Nhân dân", "Hội đồng nhân dân"), (@"Hội Đồng Nhân Dân", "Hội đồng nhân dân"),
            (@"hội đồng nhân dân", "Hội đồng nhân dân"), (@"Chính Phủ", "Chính phủ"), (@"Quốc Hội", "Quốc hội"),
            (@"Bộ Chính Trị", "Bộ Chính trị"), (@"Ban Chấp Hành", "Ban Chấp hành"), (@"Ban Thường Vụ", "Ban Thường vụ"),
            (@"Ban Bí Thư", "Ban Bí thư"), (@"Văn Phòng", "Văn phòng"), (@"Thủ Tướng", "Thủ tướng"), (@"Chủ Tịch nước", "Chủ tịch nước"),
            (@"Mặt Trận Tổ Quốc", "Mặt trận Tổ quốc"), (@"Mặt trận tổ quốc", "Mặt trận Tổ quốc"), (@"Bộ Quốc Phòng", "Bộ Quốc phòng"),
            (@"Bộ Công An", "Bộ Công an"), (@"Bộ Giáo Dục và Đào Tạo", "Bộ Giáo dục và Đào tạo"), (@"Bộ Tài Chính", "Bộ Tài chính"),
            (@"Bộ Nội Vụ", "Bộ Nội vụ"), (@"Bộ Tư Pháp", "Bộ Tư pháp"), (@"Đoàn Thanh Niên", "Đoàn Thanh niên"),
            (@"Đảng Cộng Sản Việt Nam", "Đảng Cộng sản Việt Nam"), (@"Quân Đội Nhân Dân", "Quân đội nhân dân"),
            (@"Công An Nhân Dân", "Công an nhân dân"), (@"Hiến Pháp", "Hiến pháp"), (@"Bộ Luật", "Bộ luật"),
            (@"Nghị Định", "Nghị định"), (@"Thông Tư", "Thông tư"), (@"Quyết Định", "Quyết định"), (@"Nghị Quyết", "Nghị quyết"),
            (@"việt nam", "Việt Nam"), (@"Việt nam", "Việt Nam"), (@"hà nội", "Hà Nội"), (@"Hà nội", "Hà Nội"),
            (@"thành phố Hồ Chí Minh", "Thành phố Hồ Chí Minh"), (@"Hồ chí Minh", "Hồ Chí Minh"), (@"Hồ chí minh", "Hồ Chí Minh"),
        };

        public SpellChecker(IEnumerable<string> userRuleLines = null)
        {
            // 403 cặp lỗi thường gặp (từ điển typo_corrections_map).
            var typo = SpellCheckData.TypoPairs;
            var seen = new HashSet<string>();
            for (int i = 0; i < typo.GetLength(0); i++)
            {
                var w = typo[i, 0]; var r = typo[i, 1];
                if (w != r && seen.Add(w.ToLowerInvariant()))
                    AddWordRule(w, r, "Lỗi chính tả thường gặp", ignoreCase: true);
            }
            // 50 cặp thiếu dấu / dính từ (missing_tones_corrections, confidence 1.0).
            var mt = SpellCheckData.MissingTonePairs;
            for (int i = 0; i < mt.GetLength(0); i++)
            {
                var w = mt[i, 0]; var r = mt[i, 1];
                if (w != r && seen.Add(w.ToLowerInvariant()))
                    AddWordRule(w, r, "Thiếu dấu thanh / dính từ", ignoreCase: true);
            }
            foreach (var (w, r) in ExtraMistakes.Where(x => x.wrong != x.right))
                if (seen.Add(w.ToLowerInvariant()))
                    AddWordRule(w, r, "Lỗi chính tả thường gặp", ignoreCase: true);
            foreach (var (pat, r) in Capitalization)
                AddWordRule(pat, r, "Viết hoa chưa đúng (Phụ lục II NĐ 30)", ignoreCase: false);

            // Chỉ bắt lặp hư từ — tiếng Việt có nhiều từ láy hợp lệ ("dần dần", "từ từ")
            _rules.Add(new Rule
            {
                Pattern = new Regex(@"(?<![\p{L}\p{M}])(và|của|các|những|là|được|trong|cho|với|đã|sẽ|này|theo|tại|để|thì|mà|có|không|đến|từ ngày) \1(?![\p{L}\p{M}])", RegexOptions.IgnoreCase),
                Replacement = "$1",
                Reason = "Từ bị lặp",
            });
            _rules.Add(new Rule { Pattern = new Regex(@" {2,}"), Replacement = " ", Reason = "Thừa khoảng trắng" });
            _rules.Add(new Rule { Pattern = new Regex(@" +(?=[,;:\?!]|\.(?!\.))"), Replacement = "", Reason = "Khoảng trắng trước dấu câu" });

            if (userRuleLines != null)
                foreach (var line in userRuleLines)
                {
                    var l = line.Trim();
                    if (l.Length == 0 || l.StartsWith("#")) continue;
                    var parts = l.Split('=');
                    if (parts.Length == 2 && parts[0].Trim().Length > 0)
                        AddWordRule(parts[0].Trim(), parts[1].Trim(), "Từ điển người dùng", ignoreCase: false);
                }
        }

        void AddWordRule(string wrong, string right, string reason, bool ignoreCase)
        {
            var opt = RegexOptions.CultureInvariant | (ignoreCase ? RegexOptions.IgnoreCase : RegexOptions.None);
            _rules.Add(new Rule
            {
                Pattern = new Regex(@"(?<![\p{L}\p{M}])" + Regex.Escape(wrong).Replace("\\ ", @"\s+") + @"(?![\p{L}\p{M}])", opt),
                Replacement = right,
                Reason = reason,
            });
        }

        public List<SpellHit> Check(string text, int paragraphIndex = 0)
        {
            var hits = new List<SpellHit>();
            if (string.IsNullOrEmpty(text)) return hits;
            text = text.Normalize(System.Text.NormalizationForm.FormC);

            // 1. Quét tập quy tắc từ điển và cụm từ ngữ cảnh
            foreach (var r in _rules)
                foreach (Match m in r.Pattern.Matches(text))
                {
                    if (hits.Any(h => m.Index < h.Start + h.Length && h.Start < m.Index + m.Length)) continue;
                    var sug = MatchCase(m.Value, m.Result(r.Replacement), r.Reason.StartsWith("Lỗi chính tả"));
                    if (sug == m.Value) continue;
                    hits.Add(new SpellHit
                    {
                        ParagraphIndex = paragraphIndex, Start = m.Index, Length = m.Length,
                        Found = m.Value, Suggestion = sug, Reason = r.Reason,
                    });
                }

            // 2. Vị trí đặt dấu thanh kiểu mới (QĐ 1989/QĐ-BGDĐT) trên từng âm tiết
            foreach (var t in ToneHits(text, newStyle: true, paragraphIndex))
                if (!hits.Any(h => t.Start < h.Start + h.Length && h.Start < t.Start + t.Length)) hits.Add(t);

            return hits.OrderBy(h => h.Start).ToList();
        }

        /// <summary>Các âm tiết cần đặt lại dấu thanh theo kiểu mới (hòa, thủy) hoặc kiểu cũ (hoà, thuỷ).</summary>
        public static List<SpellHit> ToneHits(string text, bool newStyle, int paragraphIndex = 0)
        {
            var hits = new List<SpellHit>();
            if (string.IsNullOrEmpty(text)) return hits;
            text = text.Normalize(System.Text.NormalizationForm.FormC);
            foreach (Match m in Regex.Matches(text, @"[\p{L}\p{M}]+"))
            {
                var fixedWord = TonePlacementValidator.PlaceTone(m.Value, newStyle);
                if (fixedWord == m.Value) continue;
                hits.Add(new SpellHit
                {
                    ParagraphIndex = paragraphIndex, Start = m.Index, Length = m.Length,
                    Found = m.Value, Suggestion = fixedWord, Reason = "Sai vị trí đặt dấu thanh",
                });
            }
            return hits;
        }

        /// <summary>Chuyển đổi vị trí dấu thanh toàn bài: newStyle=true (hòa, tòa, thủy), false (hoà, toà, thuỷ).</summary>
        public static string ConvertToneMarks(string text, bool newStyle = true) =>
            TonePlacementValidator.NormalizeAll(text, newStyle);

        /// <summary>Áp dụng mọi gợi ý vào chuỗi.</summary>
        public string Apply(string text)
        {
            foreach (var h in Check(text).OrderByDescending(h => h.Start))
                text = text.Substring(0, h.Start) + h.Suggestion + text.Substring(h.Start + h.Length);
            return text;
        }

        static string MatchCase(string found, string suggestion, bool adapt)
        {
            if (!adapt || found.Length == 0 || suggestion.Length == 0) return suggestion;
            if (TextUtil.IsAllUpper(found) && found.Length > 1) return TextUtil.Upper(suggestion);
            if (char.IsUpper(found[0])) return char.ToUpper(suggestion[0], TextUtil.Vi) + suggestion.Substring(1);
            return suggestion;
        }
    }
}
