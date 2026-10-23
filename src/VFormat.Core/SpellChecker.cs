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

        // Lỗi chính tả thường gặp (sai, đúng)
        static readonly (string wrong, string right)[] CommonMistakes =
        {
            ("sử lý", "xử lý"), ("xử lí", "xử lý"), ("sát nhập", "sáp nhập"), ("xát nhập", "sáp nhập"),
            ("chuẩn đoán", "chẩn đoán"), ("suông sẻ", "suôn sẻ"), ("xuất sắc", "xuất sắc"), ("xuất xắc", "xuất sắc"),
            ("sáng lạng", "sáng lạn"), ("xán lạng", "xán lạn"), ("bàng hoàng", "bàng hoàng"), ("dành giật", "giành giật"),
            ("tranh dành", "tranh giành"), ("dành độc lập", "giành độc lập"), ("giành dụm", "dành dụm"),
            ("dò rỉ", "rò rỉ"), ("rò hỏi", "dò hỏi"), ("chí mạng", "chí mạng"), ("trí mạng", "chí mạng"),
            ("chân thành", "chân thành"), ("trân thành", "chân thành"), ("trân trọng", "trân trọng"), ("chân trọng", "trân trọng"),
            ("cọ sát", "cọ xát"), ("sứ mệnh", "sứ mệnh"), ("xứ mệnh", "sứ mệnh"), ("xơ xuất", "sơ suất"), ("sơ xuất", "sơ suất"),
            ("xuất sứ", "xuất xứ"), ("năng xuất", "năng suất"), ("công xuất", "công suất"), ("đề suất", "đề xuất"),
            ("kiểm sát viên", "kiểm sát viên"), ("bổ xung", "bổ sung"), ("sung yếu", "xung yếu"), ("xinh hoạt", "sinh hoạt"),
            ("dữ gìn", "giữ gìn"), ("gìn dữ", "gìn giữ"), ("giữ dội", "dữ dội"), ("dang dở", "dang dở"), ("giang dở", "dang dở"),
            ("chỉn chu", "chỉn chu"), ("chỉnh chu", "chỉn chu"), ("trau dồi", "trau dồi"), ("trao dồi", "trau dồi"),
            ("lãng mạng", "lãng mạn"), ("lãng phí", "lãng phí"), ("lản phí", "lãng phí"), ("vô hình chung", "vô hình trung"),
            ("tham quang", "tham quan"), ("cập nhập", "cập nhật"),
            ("ngoan cố", "ngoan cố"), ("ngoang cố", "ngoan cố"), ("chia sẽ", "chia sẻ"),
            ("chia xẻ", "chia sẻ"), ("bảo đãm", "bảo đảm"), ("đảm bão", "đảm bảo"), ("nghành", "ngành"), ("nghiêm chỉnh", "nghiêm chỉnh"),
            ("giấy khai sinh", "giấy khai sinh"), ("xắp xếp", "sắp xếp"), ("sắp sếp", "sắp xếp"), ("sử dụng", "sử dụng"), ("xử dụng", "sử dụng"),
            ("thẩm quyền", "thẩm quyền"), ("thẫm quyền", "thẩm quyền"), ("triển khai", "triển khai"), ("chiển khai", "triển khai"),
            ("kỉ luật", "kỷ luật"), ("kĩ thuật", "kỹ thuật"), ("quí", "quý"), ("lí do", "lý do"), ("qui định", "quy định"), ("qui chế", "quy chế"),
            ("qui trình", "quy trình"), ("qui mô", "quy mô"), ("qui hoạch", "quy hoạch"), ("lí luận", "lý luận"),
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
            foreach (var (w, r) in CommonMistakes.Where(x => x.wrong != x.right))
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
            return hits.OrderBy(h => h.Start).ToList();
        }

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
