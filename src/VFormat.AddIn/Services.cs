using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using VFormat.Core;
using Word = Microsoft.Office.Interop.Word;

namespace VFormat.AddIn
{
    public class ReviewItem
    {
        public string Code, Message, Kind;
        public Severity Severity;
        public int Start = -1, End = -1;   // vị trí trong tài liệu để chọn khi bấm
        public override string ToString() => $"{Code}  {Message}";
    }

    /// <summary>Rà soát thể thức: phát hiện và đánh dấu bằng nhận xét (comment) của V-Format.</summary>
    public static class Review
    {
        public const string Author = "V-Format";
        public const string Tag = "[V-Format";

        public static List<ReviewItem> Run(Word.Application app, Word.Document doc, Settings settings, bool addComments)
        {
            var snap = WordReader.Read(doc, withFormat: true);
            var profile = settings.ResolveProfile(snap.Paragraphs);
            var cd = Classifier.Classify(snap.Paragraphs, profile.Kind);
            var issues = FormatChecker.Check(cd, profile, WordReader.ReadPage(doc));
            var items = new List<ReviewItem>();

            using (new WordBatch(app, "V-Format: Rà soát thể thức"))
            {
                ClearComments(doc, onlyVFormat: true);
                var docLevel = new StringBuilder();
                foreach (var g in issues.GroupBy(x => x.ParagraphIndex))
                {
                    if (g.Key < 0)
                    {
                        foreach (var x in g)
                        {
                            items.Add(new ReviewItem { Code = x.Code, Message = x.Message, Severity = x.Severity, Kind = "Văn bản" });
                            docLevel.AppendLine($"{x.Code}: {x.Message}");
                        }
                        continue;
                    }
                    var r = snap.WordParagraphs[g.Key].Range;
                    var kindName = FormatChecker.Describe(cd.Kinds[g.Key]);
                    foreach (var x in g)
                        items.Add(new ReviewItem { Code = x.Code, Message = x.Message, Severity = x.Severity, Kind = kindName, Start = r.Start, End = r.End });
                    if (addComments)
                        AddComment(doc, doc.Range(r.Start, Math.Max(r.Start, r.End - 1)),
                            string.Join("\n", g.Select(x => $"{x.Code}: {x.Message}")));
                }
                if (addComments && docLevel.Length > 0)
                    AddComment(doc, doc.Range(0, 0), $"Thể thức chung ({profile.FullName}):\n" + docLevel.ToString().TrimEnd());
            }
            return items;
        }

        public static void AddComment(Word.Document doc, Word.Range r, string text)
        {
            var c = doc.Comments.Add(r, Tag + "] " + text);
            try { c.Author = Author; c.Initial = "VF"; } catch { }
        }

        /// <summary>Xóa nhận xét của V-Format (hoặc tất cả nhận xét).</summary>
        public static int ClearComments(Word.Document doc, bool onlyVFormat)
        {
            int n = 0;
            for (int i = doc.Comments.Count; i >= 1; i--)
            {
                var c = doc.Comments[i];
                bool mine;
                try { mine = c.Author == Author || (c.Range.Text ?? "").StartsWith(Tag); } catch { mine = false; }
                if (!onlyVFormat || mine)
                {
                    // gỡ tô vàng do kiểm tra chính tả V-Format để lại
                    try { if (mine && (c.Range.Text ?? "").StartsWith(Tag + "] CT:")) c.Scope.HighlightColorIndex = Word.WdColorIndex.wdNoHighlight; } catch { }
                    c.Delete(); n++;
                }
            }
            return n;
        }
    }

    /// <summary>Kiểm tra chính tả offline.</summary>
    public static class Spell
    {
        static SpellChecker Build()
        {
            IEnumerable<string> user = null;
            try { if (File.Exists(Settings.UserDictionary)) user = File.ReadAllLines(Settings.UserDictionary, Encoding.UTF8); } catch { }
            return new SpellChecker(user);
        }

        public static List<ReviewItem> Run(Word.Application app, Word.Document doc, bool autoFix)
        {
            var sc = Build();
            var items = new List<ReviewItem>();
            using (new WordBatch(app, autoFix ? "V-Format: Sửa chính tả" : "V-Format: Kiểm tra chính tả"))
            {
                var paragraphs = doc.Paragraphs.Cast<Word.Paragraph>().ToList();
                for (int pi = paragraphs.Count - 1; pi >= 0; pi--)
                {
                    var pr = paragraphs[pi].Range;
                    var raw = pr.Text ?? "";
                    if (raw.Trim().Length == 0) continue;
                    var hits = sc.Check(raw, pi);
                    if (hits.Count == 0) continue;
                    bool nfc = raw.IsNormalized(NormalizationForm.FormC);
                    foreach (var h in hits.OrderByDescending(h => h.Start))
                    {
                        var r = Locate(doc, pr, h, nfc);
                        if (r == null) continue;
                        items.Add(new ReviewItem
                        {
                            Code = "CT", Kind = h.Reason, Severity = Severity.Info,
                            Message = $"\"{h.Found}\" → \"{h.Suggestion}\" ({h.Reason})", Start = r.Start, End = r.End,
                        });
                        if (autoFix) r.Text = h.Suggestion;
                        else
                        {
                            r.HighlightColorIndex = Word.WdColorIndex.wdYellow;
                            Review.AddComment(doc, r, $"CT: nên viết \"{h.Suggestion}\" — {h.Reason}");
                        }
                    }
                }
            }
            items.Reverse();
            return items;
        }

        /// <summary>
        /// Đặt lại dấu thanh toàn văn bản (kiểu mới hòa/thủy hoặc kiểu cũ hoà/thuỷ). Chỉ thay đúng từng từ,
        /// không ghi lại cả đoạn — giữ định dạng ký tự, dấu đoạn và cấu trúc bảng. Trả về số từ đã đổi.
        /// </summary>
        public static int ConvertTones(Word.Application app, Word.Document doc, bool newStyle)
        {
            int count = 0;
            using (new WordBatch(app, newStyle ? "V-Format: Dấu thanh kiểu mới" : "V-Format: Dấu thanh kiểu cũ"))
            {
                var paragraphs = doc.Paragraphs.Cast<Word.Paragraph>().ToList();
                for (int pi = paragraphs.Count - 1; pi >= 0; pi--)
                {
                    var pr = paragraphs[pi].Range;
                    var raw = pr.Text ?? "";
                    var hits = SpellChecker.ToneHits(raw, newStyle, pi);
                    if (hits.Count == 0) continue;
                    bool nfc = raw.IsNormalized(NormalizationForm.FormC);
                    foreach (var h in hits.OrderByDescending(h => h.Start))
                    {
                        var r = Locate(doc, pr, h, nfc);
                        if (r == null) continue;
                        r.Text = h.Suggestion;
                        count++;
                    }
                }
            }
            return count;
        }

        static Word.Range Locate(Word.Document doc, Word.Range para, SpellHit h, bool nfc)
        {
            if (nfc)
            {
                var r = doc.Range(para.Start + h.Start, para.Start + h.Start + h.Length);
                if (r.Text == h.Found) return r;
            }
            // văn bản tổ hợp (NFD) hoặc có trường ẩn: tìm bằng Find trong đoạn
            var fr = doc.Range(para.Start, para.End);
            var f = fr.Find;
            f.ClearFormatting();
            if (h.Found.Trim().Length == 0) return null;
            if (f.Execute(FindText: h.Found, MatchCase: true, Forward: true, Wrap: Word.WdFindWrap.wdFindStop) && fr.End <= para.End)
                return fr;
            return null;
        }
    }

    /// <summary>Chèn thành phần thể thức.</summary>
    public static class Insert
    {
        public static FormatProfile ProfileForActive(Word.Document doc, Settings s)
        {
            var snap = WordReader.Read(doc, false);
            return s.ResolveProfile(snap.Paragraphs);
        }

        public static void Header(Word.Application app, Word.Document doc, Settings s)
        {
            var p = ProfileForActive(doc, s);
            var left = new List<string>();
            var right = new List<string>();
            var ngay = s.DiaDanh + ", " + TemplateLibrary.Today();
            switch (p.Kind)
            {
                case StandardKind.Dang_HD05:
                case StandardKind.Doan_HD29:
                    left.AddRange(new[] { "TÊN CƠ QUAN CẤP TRÊN", "TÊN CƠ QUAN BAN HÀNH", p.DauSaoText, p.SoPrefix + "   -.../..." });
                    right.AddRange(new[] { p.HeaderTitle, ngay });
                    break;
                default:
                    left.AddRange(new[] { "TÊN CƠ QUAN CHỦ QUẢN", "TÊN CƠ QUAN BAN HÀNH", p.SoPrefix + "      /...-..." });
                    right.AddRange(new[] { TextNormalizer.QuocHieuText, TextNormalizer.TieuNguText, ngay });
                    break;
            }
            using (new WordBatch(app, "V-Format: Chèn khối đầu"))
            {
                var at = doc.Range(0, 0);
                at.InsertParagraphBefore();
                var table = Layout.InsertTable(doc, doc.Range(0, 0), left, right, p.HeaderLeftColMm, p.HeaderRightColMm);
                FormatInserted(doc, p, table.Range.Start, table.Range.End);
            }
        }

        public static void Signature(Word.Application app, Word.Document doc, Settings s)
        {
            var p = ProfileForActive(doc, s);
            var left = new List<string> { "Nơi nhận:", "- Như trên;", "- ...;", "- Lưu: VT, ...." };
            if (p.Kind == StandardKind.Dang_HD05) left = new List<string> { "Nơi nhận:", "- ...;", "- Lưu Văn phòng." };
            var right = new List<string> { p.ThayMatText + " TÊN TẬP THỂ", "CHỨC VỤ NGƯỜI KÝ", "", "", "", "", "Nguyễn Văn A" };
            using (new WordBatch(app, "V-Format: Chèn khối chữ ký"))
            {
                var sel = app.Selection;
                sel.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
                if (sel.Information[Word.WdInformation.wdWithInTable] is bool b && b)
                    throw new InvalidOperationException("Hãy đặt con trỏ ra ngoài bảng trước khi chèn khối chữ ký.");
                sel.TypeParagraph();
                var at = sel.Range;
                var table = Layout.InsertTable(doc, at, left, right, p.SignLeftColMm, p.SignRightColMm);
                FormatInserted(doc, p, table.Range.Start, table.Range.End);
            }
        }

        public static void NoiNhan(Word.Application app, Word.Document doc, Settings s)
        {
            var sel = app.Selection;
            sel.Collapse(Word.WdCollapseDirection.wdCollapseStart);
            int start = sel.Range.Start;
            sel.TypeText("Nơi nhận:\r- Như trên;\r- ...;\r- Lưu: VT, ....");
            var p = ProfileForActive(doc, s);
            FormatInserted(doc, p, start, sel.Range.End);
        }

        public static void Date(Word.Application app, Word.Document doc, Settings s)
        {
            var sel = app.Selection;
            int start = sel.Range.Start;
            sel.TypeText(s.DiaDanh + ", " + TemplateLibrary.Today());
            var r = doc.Range(start, sel.Range.End);
            r.Font.Italic = 1;
        }

        /// <summary>Áp định dạng chỉ cho các đoạn trong khoảng [start, end).</summary>
        static void FormatInserted(Word.Document doc, FormatProfile p, int start, int end)
        {
            var snap = WordReader.Read(doc, false);
            var cd = Classifier.Classify(snap.Paragraphs, p.Kind);
            Formatter.ApplyFormatting(doc, snap, cd, p, i => snap.Starts[i] >= start && snap.Starts[i] < end);
        }
    }

    /// <summary>Thư viện mẫu dựng sẵn + mẫu cá nhân trong %AppData%\VFormat\Templates.</summary>
    public static class Templates
    {
        public static Word.Document Create(Word.Application app, DocTemplate t, Settings s)
        {
            var doc = app.Documents.Add();
            var lines = t.Lines.Select(l => l.Replace("Hà Nội, ngày", s.DiaDanh + ", ngày"));
            doc.Content.Text = string.Join("\r", lines);
            var opts = new Settings
            {
                AutoDetect = false, Standard = t.Standard, BodySize = s.BodySize, LineSpacing = s.LineSpacing,
                FixText = true, BuildTables = true, DrawLines = true, PageNumbers = true, FormatTables = true,
            };
            Formatter.Run(app, doc, opts, t.Standard);
            return doc;
        }

        public static IEnumerable<string> UserTemplates()
        {
            try
            {
                Directory.CreateDirectory(Settings.TemplateFolder);
                return Directory.GetFiles(Settings.TemplateFolder)
                    .Where(f => f.EndsWith(".docx", StringComparison.OrdinalIgnoreCase) || f.EndsWith(".dotx", StringComparison.OrdinalIgnoreCase)
                             || f.EndsWith(".doc", StringComparison.OrdinalIgnoreCase))
                    .Where(f => !Path.GetFileName(f).StartsWith("~$"))
                    .OrderBy(f => f).ToList();
            }
            catch { return Enumerable.Empty<string>(); }
        }

        public static void OpenUser(Word.Application app, string path) => app.Documents.Add(Template: path);

        public static string SaveAsUserTemplate(Word.Document doc, string name)
        {
            if (string.IsNullOrEmpty(doc.Path))
                throw new InvalidOperationException("Hãy lưu tài liệu trước khi thêm vào kho mẫu cá nhân.");
            doc.Save();
            Directory.CreateDirectory(Settings.TemplateFolder);
            foreach (var ch in Path.GetInvalidFileNameChars()) name = name.Replace(ch, '_');
            var dest = Path.Combine(Settings.TemplateFolder, name + Path.GetExtension(doc.FullName));
            File.Copy(doc.FullName, dest, overwrite: true);
            return dest;
        }
    }
}
