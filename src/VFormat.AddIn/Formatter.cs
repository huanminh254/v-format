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
            var started = DateTime.Now;
            var result = new FormatResult();
            using (new WordBatch(app, "V-Format: Chuẩn hóa 1-Click"))
            {
                var snap = WordReader.Read(doc, withFormat: false);
                var profile = force.HasValue ? settings.ProfileFor(force.Value) : settings.ResolveProfile(snap.Paragraphs);
                result.Profile = profile;

                ApplyPageSetup(doc, profile);
                doc.Content.Font.Name = profile.FontName;

                var cd = Classifier.Classify(snap.Paragraphs, profile.Kind);
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

                if (settings.FixText)
                {
                    result.TextFixes = FixFormalText(doc, snap, cd, profile) + FixBodySpacing(doc, snap, cd);
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

        public static void ApplyPageSetup(Word.Document doc, FormatProfile p, bool? duplex = null)
        {
            foreach (Word.Section s in doc.Sections)
            {
                var ps = s.PageSetup;
                bool landscape = ps.Orientation == Word.WdOrientation.wdOrientLandscape;
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
            int lastNoiNhan = cd.Kinds.LastIndexOf(ElementKind.NoiNhanItem);
            // duyệt ngược để vị trí các đoạn phía trước không đổi
            for (int i = cd.Kinds.Count - 1; i >= 0; i--)
            {
                var kind = cd.Kinds[i];
                if (!Planner.IsFormal(kind)) continue;
                var raw = snap.Paragraphs[i].Text;
                var clean = TextUtil.Clean(raw);
                var fixedText = TextNormalizer.Fix(kind, clean, p, i == lastNoiNhan);
                if (fixedText == clean && !Regex.IsMatch(raw.TrimEnd('\r', '\a'), @"\s{2,}|^\s|\s$")) continue;
                if (fixedText.Length == 0) continue;
                int start = snap.WordParagraphs[i].Range.Start;
                int len = raw.TrimEnd('\r', '\a').Length;
                var r = doc.Range(start, start + len);
                if (r.Fields.Count > 0 || r.InlineShapes.Count > 0) continue; // không ghi đè trường/ảnh
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
                Apply(doc, para, Planner.PlanParagraph(cd.Kinds[i], info, p, inList), p.FontName);
            }
        }

        static void Apply(Word.Document doc, Word.Paragraph para, ParagraphPlan plan, string fontName)
        {
            var r = para.Range;
            var f = r.Font;
            var pf = para.Format;
            f.Name = fontName;
            if (plan.Size.HasValue) f.Size = plan.Size.Value;
            if (plan.Bold.HasValue) f.Bold = plan.Bold.Value ? 1 : 0;
            if (plan.BoldPrefixLength > 0) doc.Range(r.Start, r.Start + plan.BoldPrefixLength).Font.Bold = 1;
            if (plan.Italic.HasValue) f.Italic = plan.Italic.Value ? 1 : 0;
            if (plan.Align.HasValue)
            {
                f.Underline = Word.WdUnderline.wdUnderlineNone;
                pf.Alignment = WordReader.ToWord(plan.Align.Value);
            }
            if (plan.UnderlinePrefixLength > 0)
                doc.Range(r.Start, r.Start + plan.UnderlinePrefixLength).Font.Underline = Word.WdUnderline.wdUnderlineSingle;
            if (plan.ToUpper) r.Case = Word.WdCharacterCase.wdUpperCase;
            if (plan.ResetSideIndents) { pf.LeftIndent = 0; pf.RightIndent = 0; }
            if (plan.FirstIndentPt.HasValue) pf.FirstLineIndent = plan.FirstIndentPt.Value;
            pf.SpaceBefore = plan.SpaceBeforePt;
            pf.SpaceAfter = plan.SpaceAfterPt;
            if (plan.LineMultiple.HasValue)
            {
                pf.LineSpacingRule = Word.WdLineSpacing.wdLineSpaceMultiple;
                pf.LineSpacing = 12f * plan.LineMultiple.Value;
            }
            else pf.LineSpacingRule = Word.WdLineSpacing.wdLineSpaceSingle;
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
                    table.PreferredWidthType = Word.WdPreferredWidthType.wdPreferredWidthPercent;
                    table.PreferredWidth = 100;
                    table.AutoFitBehavior(Word.WdAutoFitBehavior.wdAutoFitWindow);
                    table.Rows.Alignment = Word.WdRowAlignment.wdAlignRowCenter;
                    if (table.Rows.Count > 2) table.Rows[1].HeadingFormat = -1;
                    table.Range.Font.Name = p.FontName;
                }
                catch { /* bảng ô gộp dọc không cho truy cập Rows[i] */ }
            }
            return n;
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
            app.ScreenUpdating = false;
            app.System.Cursor = Word.WdCursorType.wdCursorWait;
        }
        public void Dispose()
        {
            _app.ScreenUpdating = true;
            _app.System.Cursor = Word.WdCursorType.wdCursorNormal;
            if (_undo) try { _app.UndoRecord.EndCustomRecord(); } catch { }
            _app.ScreenRefresh();
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
