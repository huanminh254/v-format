using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using VFormat.Core;
using Word = Microsoft.Office.Interop.Word;

namespace VFormat.AddIn
{
    public class FormatResult
    {
        public FormatProfile Profile;
        public int Paragraphs, TextFixes, Lines, Tables;
        public bool HeaderBuilt, SignBuilt;
        public TimeSpan Elapsed;

        public override string ToString() =>
            $"Đã chuẩn hóa theo {Profile.FullName} ({Profile.Description}).\n" +
            $"• {Paragraphs} đoạn văn, {TextFixes} chỗ sửa chữ, {Tables} bảng, {Lines} đường kẻ\n" +
            (HeaderBuilt ? "• Đã dàn khối Quốc hiệu / cơ quan thành bảng 2 cột\n" : "") +
            (SignBuilt ? "• Đã dàn khối chữ ký & nơi nhận thành bảng 2 cột\n" : "") +
            $"Thời gian: {Elapsed.TotalSeconds:0.0} giây.";
    }

    /// <summary>Chuẩn hóa 1-Click.</summary>
    public static class Formatter
    {
        public static FormatResult Run(Word.Application app, Word.Document doc, Settings settings, StandardKind? force = null)
        {
            WordProtectionHelper.EnsureEditablePrintView(doc);
            var started = DateTime.Now;
            var result = new FormatResult();
            using (new WordBatch(app, "V-Format: Chuẩn hóa 1-Click"))
            {
                var snap = WordReader.Read(doc, withFormat: false);
                if (CleanupConvertedText(doc, snap)) snap = WordReader.Read(doc, withFormat: false);
                var profile = force.HasValue ? settings.ProfileFor(force.Value) : settings.ResolveProfile(snap.Paragraphs);
                result.Profile = profile;

                var cd = Classifier.Classify(snap.Paragraphs, profile.Kind);
                ApplyPageSetup(doc, profile, landscapeAllowed: Planner.HasContentTable(cd));
                using (new SymbolFontScope(doc))
                {
                    doc.Content.Font.Name = profile.FontName;
                    try { doc.Content.Font.NameAscii = profile.FontName; } catch { }
                    try { doc.Content.Font.NameOther = profile.FontName; } catch { }
                }

                if (settings.BuildTables)
                {
                    if (Layout.BuildHeaderTable(doc, snap, cd, profile))
                    {
                        result.HeaderBuilt = true;
                        snap = WordReader.Read(doc, false); cd = Classifier.Classify(snap.Paragraphs, profile.Kind);
                    }
                    if (Layout.BuildSignTable(doc, snap, cd, profile))
                    {
                        result.SignBuilt = true;
                        snap = WordReader.Read(doc, false); cd = Classifier.Classify(snap.Paragraphs, profile.Kind);
                    }
                }

                if (MoveCongVanSubject(doc, snap, cd))
                {
                    snap = WordReader.Read(doc, false); cd = Classifier.Classify(snap.Paragraphs, profile.Kind);
                }
                if (MoveConsolidatedLabel(doc, snap, cd, profile))
                {
                    snap = WordReader.Read(doc, false); cd = Classifier.Classify(snap.Paragraphs, profile.Kind);
                }

                if (settings.FixText)
                {
                    result.TextFixes = FixFormalText(doc, snap, cd, profile) + FixBodySpacing(doc, snap, cd);
                    snap = WordReader.Read(doc, false); cd = Classifier.Classify(snap.Paragraphs, profile.Kind);
                }

                if (InsertTitleDashes(doc, snap, cd, profile))
                {
                    snap = WordReader.Read(doc, false); cd = Classifier.Classify(snap.Paragraphs, profile.Kind);
                }

                if (NormalizeBlankLines(doc, snap, cd))
                {
                    snap = WordReader.Read(doc, false); cd = Classifier.Classify(snap.Paragraphs, profile.Kind);
                }

                ApplyFormatting(doc, snap, cd, profile);
                result.Paragraphs = snap.Paragraphs.Count;
                if (settings.FormatTables) result.Tables = FormatTables(doc, snap, cd, profile);
                if (settings.PageNumbers) PageNumbers.Apply(doc, profile);

                if (settings.DrawLines)
                {
                    doc.ActiveWindow.View.Type = Word.WdViewType.wdPrintView;
                    doc.Repaginate();
                    result.Lines = Layout.DrawLines(doc, snap, cd, profile);
                }
            }
            result.Elapsed = DateTime.Now - started;
            return result;
        }

        /// <summary>
        /// Văn bản chuyển từ PDF / gõ Shift+Enter: tách khối đầu bị nhồi trong một đoạn, nối hoặc tách từng ngắt dòng cứng
        /// (dòng kết thúc bằng ngắt dòng trong đoạn căn đều bị giãn chữ), bỏ số trang PDF sót lại, bỏ ngắt trang giữa các section.
        /// </summary>
        static bool CleanupConvertedText(Word.Document doc, WordSnapshot snap)
        {
            bool converted = ConvertedText.LooksConverted(snap.Paragraphs);
            bool changed = false;
            var pageNumbers = new HashSet<int>();
            if (converted)
                for (int s = 2; s <= doc.Sections.Count; s++)
                {
                    int secStart = doc.Sections[s].Range.Start;
                    int i = snap.Starts.FindIndex(x => x >= secStart);
                    while (i >= 0 && i < snap.Paragraphs.Count && !snap.Paragraphs[i].InTable && TextUtil.Clean(snap.Paragraphs[i].Text).Length == 0) i++;
                    if (i >= 0 && i < snap.Paragraphs.Count && !snap.Paragraphs[i].InTable && ConvertedText.IsPageNumberArtifact(snap.Paragraphs[i].Text))
                        pageNumbers.Add(i);
                }

            for (int i = snap.Paragraphs.Count - 1; i >= 0; i--)
            {
                var info = snap.Paragraphs[i];
                if (info.InTable) continue;
                var raw = info.Text.TrimEnd('\r', '\a');
                int start = snap.Starts[i];
                var r = doc.Range(start, start + raw.Length);
                if (r.Fields.Count > 0 || r.InlineShapes.Count > 0) continue;   // vị trí ký tự lệch với Range.Text
                if (pageNumbers.Contains(i)) { snap.WordParagraphs[i].Range.Delete(); changed = true; continue; }

                var header = i < 20 ? ConvertedText.SplitMergedHeader(raw) : null;
                if (header != null)
                {
                    r.Text = string.Join("\r", header);
                    r.Font.Superscript = 0; r.Font.Subscript = 0;   // số chú thích "¹CỤC" đứng đầu đoạn cũ
                    changed = true;
                    continue;
                }

                if (raw.IndexOf(ConvertedText.LineBreak) >= 0)
                {
                    var plan = ConvertedText.Plan(raw);
                    var pos = new List<int>();
                    for (int c = 0; c < raw.Length; c++) if (raw[c] == ConvertedText.LineBreak) pos.Add(c);
                    var action = new Dictionary<int, LineBreakAction>();
                    for (int k = 0; k < pos.Count; k++) action[pos[k]] = plan[k];
                    // từ cuối lên: vị trí phía trước không đổi
                    for (int k = pos.Count - 1; k >= 0; k--)
                    {
                        int c = pos[k], len = 1;
                        bool spaced = c > 0 && raw[c - 1] == ' ' || c + 1 < raw.Length && raw[c + 1] == ' ';
                        if (plan[k] == LineBreakAction.Split)
                        {
                            // bỏ tab / khoảng trắng đầu đoạn mới (kể cả sau các ngắt dòng đã bỏ — chúng đã bị xóa)
                            for (int j = c + 1; j < raw.Length; j++)
                            {
                                if (raw[j] == ' ' || raw[j] == '\t') len++;
                                else if (raw[j] != ConvertedText.LineBreak || action[j] != LineBreakAction.Remove) break;
                            }
                        }
                        doc.Range(start + c, start + c + len).Text =
                            plan[k] == LineBreakAction.Split ? "\r" : plan[k] == LineBreakAction.Join && !spaced ? " " : "";
                    }
                    changed = true;
                }
                // văn bản chuyển từ PDF: tab / khoảng trắng đầu đoạn dựng vị trí chữ, chồng lên thụt đầu dòng chuẩn
                int lead = converted ? ConvertedText.LeadingBlank(raw, 0) : 0;
                if (lead > 0 && lead < raw.Length) { doc.Range(start, start + lead).Text = ""; changed = true; }
            }

            if (converted)
                for (int s = 2; s <= doc.Sections.Count; s++)
                    try { doc.Sections[s].PageSetup.SectionStart = Word.WdSectionStart.wdSectionContinuous; } catch { }
            return changed;
        }

        /// <param name="landscapeAllowed">false: trang ngang được xoay về dọc (văn bản không có bảng nội dung)</param>
        public static void ApplyPageSetup(Word.Document doc, FormatProfile p, bool? duplex = null, bool landscapeAllowed = true)
        {
            foreach (Word.Section s in doc.Sections)
            {
                var ps = s.PageSetup;
                bool landscape = ps.Orientation == Word.WdOrientation.wdOrientLandscape;
                if (landscape && !landscapeAllowed) { ps.Orientation = Word.WdOrientation.wdOrientPortrait; landscape = false; }
                ps.PageWidth = (landscape ? p.PageHeightMm : p.PageWidthMm) * WordReader.PtPerMm;
                ps.PageHeight = (landscape ? p.PageWidthMm : p.PageHeightMm) * WordReader.PtPerMm;
                ps.TopMargin = p.MarginTopMm * WordReader.PtPerMm;
                ps.BottomMargin = p.MarginBottomMm * WordReader.PtPerMm;
                ps.LeftMargin = p.MarginLeftMm * WordReader.PtPerMm;
                ps.RightMargin = p.MarginRightMm * WordReader.PtPerMm;
                ps.Gutter = 0;
                ps.HeaderDistance = 10 * WordReader.PtPerMm;
                ps.FooterDistance = 10 * WordReader.PtPerMm;
                if (duplex.HasValue) ps.MirrorMargins = duplex.Value ? -1 : 0;
                else ps.MirrorMargins = DocState.GetDuplex(doc) ? -1 : 0;
            }
        }

        /// <summary>Sửa chữ của các thành phần thể thức (Quốc hiệu, tiêu ngữ, số ký hiệu, ngày tháng, nơi nhận...).</summary>
        static int FixFormalText(Word.Document doc, WordSnapshot snap, ClassifiedDocument cd, FormatProfile p)
        {
            int fixes = 0;
            // duyệt ngược để vị trí các đoạn phía trước không đổi
            for (int i = cd.Kinds.Count - 1; i >= 0; i--)
            {
                var kind = cd.Kinds[i];
                if (!Planner.NeedsTextFix(cd, i, p)) continue;
                if (snap.Paragraphs[i].InTable && cd.Kinds[i] == ElementKind.ThongTinDau) continue;
                var raw = snap.Paragraphs[i].Text;
                var clean = TextUtil.Clean(raw);
                var fixedText = Planner.FixText(cd, i, p, clean);
                if (fixedText == clean && !Regex.IsMatch(raw.TrimEnd('\r', '\a'), @"\s{2,}|^\s|\s$")) continue;
                if (fixedText.Length == 0) continue;
                int start = snap.WordParagraphs[i].Range.Start;
                int len = raw.TrimEnd('\r', '\a').Length;
                var r = doc.Range(start, start + len);
                if (r.Fields.Count > 0 || r.InlineShapes.Count > 0 || r.Footnotes.Count > 0) continue; // không ghi đè trường/ảnh/chú thích
                r.Text = fixedText;
                fixes++;
            }
            return fixes;
        }

        /// <summary>Khoảng trắng thừa, khoảng trắng trước dấu câu trong phần nội dung (giữ nguyên định dạng ký tự).</summary>
        static int FixBodySpacing(Word.Document doc, WordSnapshot snap, ClassifiedDocument cd)
        {
            var (first, last) = Planner.BodyRange(cd);
            if (first < 0) return 0;
            int count = 0;
            var rx = new[]
            {
                ("[ ]{2,}", " "),
                ("[ ]{1,}([,;:\\?\\!\\)])", "\\1"),
                ("[ ]{1,}(.)([!.])", "\\1\\2"),   // dấu chấm đơn; giữ " ..." (chỗ để trống)
                ("([\\(])[ ]{1,}", "\\1"),        // Word từ chối "(\()": dấu ngoặc phải nằm trong [ ]
                ("([0-9])[ ]{1,}/", "\\1/"),        // "số 15 /KH-VP"
            };
            foreach (var (find, repl) in rx)
            {
                var range = doc.Range(snap.WordParagraphs[first].Range.Start, snap.WordParagraphs[last].Range.End);
                var f = range.Find;
                f.ClearFormatting(); f.Replacement.ClearFormatting();
                bool found = f.Execute(FindText: find, MatchWildcards: true, ReplaceWith: repl,
                    Replace: Word.WdReplace.wdReplaceAll, Wrap: Word.WdFindWrap.wdFindStop, Forward: true);
                if (found) count++;
            }
            return count;
        }

        /// <summary>Công văn: đưa "V/v ..." vào ô trái khối đầu dưới số, ký hiệu; bỏ dòng nhãn "công văn".</summary>
        static bool MoveCongVanSubject(Word.Document doc, WordSnapshot snap, ClassifiedDocument cd)
        {
            var (from, after) = Planner.PlanCongVanSubjectMove(cd);
            int stray = Planner.StrayCongVanLabel(cd);
            if (from < 0 && stray < 0) return false;
            string subject = from >= 0 ? TextUtil.Clean(snap.Paragraphs[from].Text) : null;
            // xóa từ dưới lên rồi mới chèn vào bảng phía trên
            foreach (var i in new[] { from, stray }.Where(x => x >= 0).OrderByDescending(x => x))
                snap.WordParagraphs[i].Range.Delete();
            if (subject != null)
            {
                var so = snap.WordParagraphs[after].Range;
                doc.Range(so.End - 1, so.End - 1).InsertAfter("\r" + subject);
            }
            return true;
        }

        /// <summary>Văn bản hợp nhất: đưa "VĂN BẢN HỢP NHẤT" từ giữa trang vào ô trái khối đầu, trên số, ký hiệu.</summary>
        static bool MoveConsolidatedLabel(Word.Document doc, WordSnapshot snap, ClassifiedDocument cd, FormatProfile p)
        {
            var (from, before) = Planner.PlanConsolidatedLabelMove(cd, p);
            if (from < 0) return false;
            var label = TextUtil.Clean(snap.Paragraphs[from].Text);
            int start = snap.WordParagraphs[before].Range.Start;   // đoạn số ký hiệu nằm trước: xóa sau không làm lệch
            snap.WordParagraphs[from].Range.Delete();
            doc.Range(start, start).InsertBefore(label + "\r");
            return true;
        }

        /// <summary>Đảng, Đoàn: chèn dòng "-----" dưới tên loại, trích yếu nếu chưa có.</summary>
        static bool InsertTitleDashes(Word.Document doc, WordSnapshot snap, ClassifiedDocument cd, FormatProfile p)
        {
            int after = Planner.PlanTitleDashes(cd, p);
            if (after < 0) return false;
            var r = snap.WordParagraphs[after].Range;
            doc.Range(r.End - 1, r.End - 1).InsertAfter("\r" + p.TitleDashes);
            return true;
        }

        public const int SignatureBlankLines = 4;

        /// <summary>
        /// Bỏ dòng trống thừa (trước tên loại, nhiều dòng trống liền nhau, cuối văn bản)
        /// và chừa đủ chỗ ký, đóng dấu giữa chức vụ và họ tên người ký.
        /// </summary>
        static bool NormalizeBlankLines(Word.Document doc, WordSnapshot snap, ClassifiedDocument cd)
        {
            var k = cd.Kinds;
            int n = k.Count;
            // dấu cuối hàng của bảng cũng là một "đoạn" rỗng ngoài ô: không đụng tới
            bool FreeEmpty(int i) => i >= 0 && i < n && k[i] == ElementKind.Empty && !snap.Paragraphs[i].InTable
                                     && !(bool)snap.WordParagraphs[i].Range.Information[Word.WdInformation.wdWithInTable];
            int lastText = k.FindLastIndex(x => x != ElementKind.Empty);

            var ops = new List<(int index, int insert)>(); // insert = 0: xóa đoạn trống; > 0: chèn n dòng trống trước đoạn
            for (int i = 0; i < n - 1; i++)
            {
                if (!FreeEmpty(i)) continue;
                int next = i + 1;
                while (FreeEmpty(next)) next++;
                bool beforeTitle = next < n && k[next] == ElementKind.TenLoai;
                if (beforeTitle || FreeEmpty(i + 1) || i > lastText) ops.Add((i, 0));
            }

            int name = k.LastIndexOf(ElementKind.HoTenKy);
            if (name > 0)
            {
                int j = name - 1, blanks = 0;
                while (j >= 0 && k[j] == ElementKind.Empty) { blanks++; j--; }
                if (j >= 0 && k[j] == ElementKind.ChucVuKy && blanks < SignatureBlankLines)
                    ops.Add((name, SignatureBlankLines - blanks));
            }

            // từ cuối lên để vị trí các đoạn phía trước không đổi
            foreach (var (index, insert) in ops.OrderByDescending(o => o.index))
            {
                try
                {
                    var r = snap.WordParagraphs[index].Range;
                    if (insert == 0) r.Delete();
                    else doc.Range(r.Start, r.Start).InsertBefore(new string('\r', insert)); // đoạn cuối ô bảng: InsertParagraphBefore lỗi
                }
                catch { /* đoạn trống ngăn hai bảng: Word không cho xóa */ }
            }
            return ops.Count > 0;
        }

        /// <summary>Áp phông, cỡ, kiểu chữ, căn lề, giãn dòng theo thành phần.</summary>
        public static void ApplyFormatting(Word.Document doc, WordSnapshot snap, ClassifiedDocument cd, FormatProfile p,
            Func<int, bool> filter = null)
        {
            for (int i = 0; i < cd.Kinds.Count; i++)
            {
                if (filter != null && !filter(i)) continue;
                var para = snap.WordParagraphs[i];
                var r = para.Range;
                bool inList = r.ListFormat.ListType != Word.WdListType.wdListNoNumbering;
                var info = snap.Paragraphs[i];
                if (cd.Kinds[i] == ElementKind.TableText && info.FontSize <= 0)
                    info.FontSize = r.Font.Size >= 9999999 ? 0 : r.Font.Size;
                if (cd.Kinds[i] == ElementKind.Body || cd.Kinds[i] == ElementKind.TableText)
                {
                    // đậm / nghiêng phủ cả đoạn (snapshot đọc không kèm định dạng)
                    int b = r.Font.Bold, it = r.Font.Italic;
                    info.Bold = b == 9999999 ? -1 : b != 0 ? 1 : 0;
                    info.Italic = it == 9999999 ? -1 : it != 0 ? 1 : 0;
                }
                Apply(doc, para, Planner.PlanParagraph(cd, i, p, inList), p.FontName);
            }
        }

        static void Apply(Word.Document doc, Word.Paragraph para, ParagraphPlan plan, string fontName)
        {
            var r = para.Range;
            var f = r.Font;
            var pf = para.Format;
            f.Name = fontName;
            if (plan.ResetColor) f.Color = Word.WdColor.wdColorAutomatic;
            if (plan.ClearUnderline) f.Underline = Word.WdUnderline.wdUnderlineNone;
            if (plan.Size.HasValue) f.Size = plan.Size.Value;
            if (plan.Bold.HasValue) f.Bold = plan.Bold.Value ? 1 : 0;
            if (plan.BoldPrefixLength > 0) doc.Range(r.Start, r.Start + plan.BoldPrefixLength).Font.Bold = 1;
            if (plan.Italic.HasValue) f.Italic = plan.Italic.Value ? 1 : 0;
            if (plan.ItalicPrefixLength > 0) doc.Range(r.Start, r.Start + plan.ItalicPrefixLength).Font.Italic = 1;
            if (plan.Align.HasValue)
            {
                f.Underline = Word.WdUnderline.wdUnderlineNone;
                pf.Alignment = WordReader.ToWord(plan.Align.Value);
            }
            if (plan.UnderlinePrefixLength > 0)
                doc.Range(r.Start, r.Start + plan.UnderlinePrefixLength).Font.Underline = Word.WdUnderline.wdUnderlineSingle;
            if (plan.ToUpper) r.Case = Word.WdCharacterCase.wdUpperCase;
            if (plan.ResetSideIndents) { pf.LeftIndent = 0; pf.RightIndent = 0; }
            if (plan.LeftIndentPt.HasValue) { pf.LeftIndent = plan.LeftIndentPt.Value; pf.RightIndent = 0; }
            if (plan.FirstIndentPt.HasValue) pf.FirstLineIndent = plan.FirstIndentPt.Value;
            pf.SpaceBefore = plan.SpaceBeforePt;
            pf.SpaceAfter = plan.SpaceAfterPt;
            if (plan.LineExactPt.HasValue)
            {
                pf.LineSpacingRule = Word.WdLineSpacing.wdLineSpaceExactly;
                pf.LineSpacing = plan.LineExactPt.Value;
            }
            else if (plan.LineMultiple.HasValue)
            {
                pf.LineSpacingRule = Word.WdLineSpacing.wdLineSpaceMultiple;
                pf.LineSpacing = 12f * plan.LineMultiple.Value;
            }
            else pf.LineSpacingRule = Word.WdLineSpacing.wdLineSpaceSingle;
            if (plan.BorderBottom)
            {
                var b = pf.Borders[Word.WdBorderType.wdBorderBottom];
                b.LineStyle = Word.WdLineStyle.wdLineStyleSingle;
                b.LineWidth = Word.WdLineWidth.wdLineWidth075pt;
                b.Color = Word.WdColor.wdColorAutomatic;
            }
            pf.WidowControl = -1;
            pf.KeepWithNext = plan.KeepWithNext ? -1 : 0;
        }

        /// <summary>Bảng nội dung: không tràn lề, tự lặp hàng tiêu đề; bảng khối đầu/chữ ký: bỏ viền.</summary>
        static int FormatTables(Word.Document doc, WordSnapshot snap, ClassifiedDocument cd, FormatProfile p)
        {
            var (headerTables, signTables) = Planner.FormalTables(cd);
            int n = 0;
            for (int t = 1; t <= doc.Tables.Count; t++)
            {
                var table = doc.Tables[t];
                n++;
                if (headerTables.Contains(t)) { Layout.NormalizeFormalTable(table, p.HeaderLeftColMm, p.HeaderRightColMm); continue; }
                if (signTables.Contains(t)) { Layout.NormalizeFormalTable(table, p.SignLeftColMm, p.SignRightColMm); continue; }
                try
                {
                    table.Rows.LeftIndent = 0;
                    var ps = table.Range.Sections[1].PageSetup;
                    float availMm = (ps.PageWidth - ps.LeftMargin - ps.RightMargin) / WordReader.PtPerMm;
                    var widths = Planner.PlanColumnWidths(cd, t, availMm);
                    if (widths != null)
                    {
                        // độ rộng cột theo nội dung (dùng chung với engine .docx), đủ khổ vùng chữ
                        table.AllowAutoFit = false;
                        table.PreferredWidthType = Word.WdPreferredWidthType.wdPreferredWidthPoints;
                        table.PreferredWidth = widths.Sum() * WordReader.PtPerMm;
                        foreach (Word.Cell cell in table.Range.Cells)
                        {
                            cell.PreferredWidthType = Word.WdPreferredWidthType.wdPreferredWidthPoints;
                            cell.Width = widths[cell.ColumnIndex - 1] * WordReader.PtPerMm;
                        }
                    }
                    else
                    {
                        // ô gộp: bỏ độ rộng cố định, co theo nội dung rồi giãn đủ khổ
                        table.PreferredWidthType = Word.WdPreferredWidthType.wdPreferredWidthPercent;
                        table.PreferredWidth = 100;
                        table.AllowAutoFit = true;
                        table.Range.Cells.PreferredWidthType = Word.WdPreferredWidthType.wdPreferredWidthAuto;
                        table.AutoFitBehavior(Word.WdAutoFitBehavior.wdAutoFitContent);
                        table.AutoFitBehavior(Word.WdAutoFitBehavior.wdAutoFitWindow);
                    }
                    table.Rows.Alignment = Word.WdRowAlignment.wdAlignRowCenter;
                    if (table.Rows.Count > 2) table.Rows[1].HeadingFormat = -1;
                    table.Range.Font.Name = p.FontName;
                    FormatDataCells(table);
                }
                catch { /* bảng ô gộp dọc không cho truy cập Rows[i] */ }
            }
            return n;
        }

        static readonly Regex RxNumberCell = new Regex(@"^[\d\s\.,%+\-–/()]+$");

        /// <summary>Bảng số liệu: hàng tiêu đề in đậm, căn giữa; ô chỉ chứa số căn giữa.</summary>
        static void FormatDataCells(Word.Table table)
        {
            bool headerRow = table.Rows.Count > 1;
            foreach (Word.Cell cell in table.Range.Cells)
            {
                var text = TextUtil.Clean(cell.Range.Text);
                if (cell.RowIndex == 1 && text.Length > 0 && (text.Length > 60 || RxNumberCell.IsMatch(text))) headerRow = false;
            }
            foreach (Word.Cell cell in table.Range.Cells)
            {
                var text = TextUtil.Clean(cell.Range.Text);
                bool isHeader = headerRow && cell.RowIndex == 1;
                if (isHeader) cell.Range.Font.Bold = 1;
                if (isHeader || text.Length > 0 && RxNumberCell.IsMatch(text))
                {
                    cell.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                    cell.VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter;
                }
            }
        }

        public static void FormatHeaderOnly(Word.Application app, Word.Document doc, Settings settings)
        {
            WordProtectionHelper.EnsureEditablePrintView(doc);
            using (new WordBatch(app, "V-Format: Chuẩn hóa khối Tiêu đề"))
            {
                var snap = WordReader.Read(doc, withFormat: false);
                var profile = settings.ResolveProfile(snap.Paragraphs);
                var cd = Classifier.Classify(snap.Paragraphs, profile.Kind);
                if (settings.BuildTables)
                {
                    if (Layout.BuildHeaderTable(doc, snap, cd, profile))
                    {
                        snap = WordReader.Read(doc, false); cd = Classifier.Classify(snap.Paragraphs, profile.Kind);
                    }
                }
                ApplyFormatting(doc, snap, cd, profile, i =>
                {
                    var k = cd.Kinds[i];
                    return k == ElementKind.QuocHieu || k == ElementKind.TieuNgu ||
                           k == ElementKind.TieuDeDang || k == ElementKind.TieuDeDoan ||
                           k == ElementKind.CoQuanChuQuan || k == ElementKind.CoQuanBanHanh ||
                           k == ElementKind.DauSao || k == ElementKind.SoKyHieu ||
                           k == ElementKind.TrichYeuCongVan || k == ElementKind.DiaDanhNgayThang;
                });
                if (settings.DrawLines)
                {
                    doc.Repaginate();
                    Layout.DrawLines(doc, snap, cd, profile);
                }
            }
        }

        public static void FormatSignOnly(Word.Application app, Word.Document doc, Settings settings)
        {
            WordProtectionHelper.EnsureEditablePrintView(doc);
            using (new WordBatch(app, "V-Format: Chuẩn hóa khối Chữ ký"))
            {
                var snap = WordReader.Read(doc, withFormat: false);
                var profile = settings.ResolveProfile(snap.Paragraphs);
                var cd = Classifier.Classify(snap.Paragraphs, profile.Kind);
                if (settings.BuildTables)
                {
                    if (Layout.BuildSignTable(doc, snap, cd, profile))
                    {
                        snap = WordReader.Read(doc, false); cd = Classifier.Classify(snap.Paragraphs, profile.Kind);
                    }
                }
                ApplyFormatting(doc, snap, cd, profile, i =>
                {
                    var k = cd.Kinds[i];
                    return k == ElementKind.ChucVuKy || k == ElementKind.HoTenKy ||
                           k == ElementKind.NoiNhanTitle || k == ElementKind.NoiNhanItem;
                });
                if (settings.DrawLines)
                {
                    doc.Repaginate();
                    Layout.DrawLines(doc, snap, cd, profile);
                }
            }
        }

        public static void FormatBodyOnly(Word.Application app, Word.Document doc, Settings settings)
        {
            WordProtectionHelper.EnsureEditablePrintView(doc);
            using (new WordBatch(app, "V-Format: Chuẩn hóa phần Thân bài"))
            {
                var snap = WordReader.Read(doc, withFormat: false);
                var profile = settings.ResolveProfile(snap.Paragraphs);
                var cd = Classifier.Classify(snap.Paragraphs, profile.Kind);
                ApplyFormatting(doc, snap, cd, profile, i =>
                {
                    var k = cd.Kinds[i];
                    return k != ElementKind.QuocHieu && k != ElementKind.TieuNgu &&
                           k != ElementKind.TieuDeDang && k != ElementKind.TieuDeDoan &&
                           k != ElementKind.CoQuanChuQuan && k != ElementKind.CoQuanBanHanh &&
                           k != ElementKind.DauSao && k != ElementKind.SoKyHieu &&
                           k != ElementKind.TrichYeuCongVan && k != ElementKind.DiaDanhNgayThang &&
                           k != ElementKind.ChucVuKy && k != ElementKind.HoTenKy &&
                           k != ElementKind.NoiNhanTitle && k != ElementKind.NoiNhanItem;
                });
                if (settings.FormatTables) FormatTables(doc, snap, cd, profile);
            }
        }
    }

    /// <summary>Gom các thay đổi thành một bước Undo và tắt cập nhật màn hình.</summary>
    public sealed class WordBatch : IDisposable
    {
        readonly Word.Application _app;
        readonly bool _undo;
        public WordBatch(Word.Application app, string name)
        {
            _app = app;
            try { app.UndoRecord.StartCustomRecord(name); _undo = true; } catch { }
            try { app.ScreenUpdating = false; } catch { }
            try { app.DisplayAlerts = Word.WdAlertLevel.wdAlertsNone; } catch { }
            try { app.System.Cursor = Word.WdCursorType.wdCursorWait; } catch { }
        }
        public void Dispose()
        {
            try { _app.ScreenUpdating = true; } catch { }
            try { _app.DisplayAlerts = Word.WdAlertLevel.wdAlertsAll; } catch { }
            try { _app.System.Cursor = Word.WdCursorType.wdCursorNormal; } catch { }
            if (_undo) try { _app.UndoRecord.EndCustomRecord(); } catch { }
            try { _app.ScreenRefresh(); } catch { }
            GC.Collect();
            GC.WaitForPendingFinalizers();
        }
    }

    /// <summary>Trạng thái lưu riêng cho từng tài liệu (in một mặt / hai mặt).</summary>
    public static class DocState
    {
        const string DuplexVar = "VFormat_Duplex";

        public static bool GetDuplex(Word.Document doc)
        {
            foreach (Word.Variable v in doc.Variables) if (v.Name == DuplexVar) return v.Value == "1";
            return doc.Sections[1].PageSetup.MirrorMargins != 0;
        }

        public static void SetDuplex(Word.Document doc, bool on)
        {
            foreach (Word.Variable v in doc.Variables) if (v.Name == DuplexVar) { v.Value = on ? "1" : "0"; return; }
            doc.Variables.Add(DuplexVar, on ? "1" : "0");
        }
    }
}
