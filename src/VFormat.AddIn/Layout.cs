using System;
using System.Collections.Generic;
using System.Linq;
using VFormat.Core;
using Word = Microsoft.Office.Interop.Word;

namespace VFormat.AddIn
{
    /// <summary>Dàn khối đầu / khối chữ ký thành bảng 2 cột không viền và vẽ đường kẻ ngang theo quy định.</summary>
    public static class Layout
    {
        public const string ShapePrefix = "VFormat_Line_";

        /// <summary>Khối đầu đang là các đoạn rời → chuyển thành bảng 2 cột. Trả về true nếu đã dàn lại.</summary>
        public static bool BuildHeaderTable(Word.Document doc, WordSnapshot snap, ClassifiedDocument cd, FormatProfile p) =>
            Apply(doc, snap, Planner.PlanHeader(cd, p));

        /// <summary>Khối chữ ký + nơi nhận đang là các đoạn rời → bảng 2 cột (trái: nơi nhận, phải: chữ ký).</summary>
        public static bool BuildSignTable(Word.Document doc, WordSnapshot snap, ClassifiedDocument cd, FormatProfile p) =>
            Apply(doc, snap, Planner.PlanSign(cd, p));

        static bool Apply(Word.Document doc, WordSnapshot snap, TablePlan plan)
        {
            if (plan == null) return false;
            ReplaceWithTable(doc, snap.WordParagraphs[plan.First].Range.Start, snap.WordParagraphs[plan.Last].Range.End,
                plan.Left, plan.Right, plan.LeftMm, plan.RightMm);
            return true;
        }

        public static Word.Table ReplaceWithTable(Word.Document doc, int start, int end, IList<string> left, IList<string> right,
            float leftMm, float rightMm)
        {
            var region = doc.Range(start, end);
            region.Delete();
            var at = doc.Range(start, start);
            at.InsertParagraphBefore();
            return InsertTable(doc, doc.Range(start, start), left, right, leftMm, rightMm);
        }

        public static Word.Table InsertTable(Word.Document doc, Word.Range at, IList<string> left, IList<string> right,
            float leftMm, float rightMm)
        {
            var table = doc.Tables.Add(at, 1, 2);
            table.Borders.Enable = 0;
            table.AllowAutoFit = false;
            table.TopPadding = 0; table.BottomPadding = 0;
            table.LeftPadding = 0; table.RightPadding = 0;
            table.Rows.LeftIndent = 0;
            table.Columns[1].SetWidth(leftMm * WordReader.PtPerMm, Word.WdRulerStyle.wdAdjustNone);
            table.Columns[2].SetWidth(rightMm * WordReader.PtPerMm, Word.WdRulerStyle.wdAdjustNone);
            table.Cell(1, 1).Range.Text = string.Join("\r", left);
            table.Cell(1, 2).Range.Text = string.Join("\r", right);
            foreach (Word.Paragraph cp in table.Range.Paragraphs)
            {
                cp.Format.FirstLineIndent = 0; cp.Format.LeftIndent = 0; cp.Format.RightIndent = 0;
                cp.Format.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
            }
            table.Rows.AllowBreakAcrossPages = 0;
            return table;
        }

        /// <summary>Bảng khối đầu / chữ ký: bỏ viền, đặt độ rộng cột.</summary>
        public static void NormalizeFormalTable(Word.Table t, float leftMm, float rightMm)
        {
            t.Borders.Enable = 0;
            try
            {
                if (t.Columns.Count == 2)
                {
                    t.AllowAutoFit = false;
                    t.Columns[1].SetWidth(leftMm * WordReader.PtPerMm, Word.WdRulerStyle.wdAdjustNone);
                    t.Columns[2].SetWidth(rightMm * WordReader.PtPerMm, Word.WdRulerStyle.wdAdjustNone);
                }
            }
            catch { /* bảng có ô gộp: giữ nguyên độ rộng */ }
        }

        // ---------------- Đường kẻ ----------------

        public static void RemoveLines(Word.Document doc)
        {
            for (int i = doc.Shapes.Count; i >= 1; i--)
            {
                var s = doc.Shapes[i];
                if (s.Name != null && s.Name.StartsWith(ShapePrefix)) s.Delete();
            }
        }

        public static int DrawLines(Word.Document doc, WordSnapshot snap, ClassifiedDocument cd, FormatProfile p)
        {
            RemoveLines(doc);
            int n = 0;
            foreach (var (i, ratio) in Planner.PlanLines(cd, p))
                if (DrawUnder(doc, snap.WordParagraphs[i], snap.Paragraphs[i].Text, ratio, n)) n++;
            return n;
        }

        static float X(Word.Range r) => Convert.ToSingle(r.Information[Word.WdInformation.wdHorizontalPositionRelativeToPage]);
        static float Y(Word.Range r) => Convert.ToSingle(r.Information[Word.WdInformation.wdVerticalPositionRelativeToPage]);

        static bool DrawUnder(Word.Document doc, Word.Paragraph para, string rawText, float ratio, int n)
        {
            var pr = para.Range;
            int len = rawText.TrimEnd('\r', '\a', '\v', ' ').Length;
            if (len == 0) return false;
            int start = pr.Start, end = pr.Start + len;
            var endR = doc.Range(end, end);
            float yEnd = Y(endR);
            if (yEnd < 0) return false;

            // đầu dòng cuối cùng (đoạn nhiều dòng): tìm nhị phân ký tự đầu tiên cùng dòng với cuối đoạn
            int lo = start, hi = end;
            while (lo < hi)
            {
                int mid = (lo + hi) / 2;
                if (Math.Abs(Y(doc.Range(mid, mid)) - yEnd) < 1) hi = mid; else lo = mid + 1;
            }
            float x1 = X(doc.Range(lo, lo)), x2 = X(endR);
            if (x2 <= x1) return false;

            float size = pr.Font.Size >= 9999999 ? 13 : pr.Font.Size;
            float y = yEnd + size * 1.2f + 1f;
            float w = (x2 - x1) * ratio, left = (x1 + x2) / 2 - w / 2;

            var shape = doc.Shapes.AddLine(left, y, left + w, y, pr);
            shape.Name = ShapePrefix + n;
            shape.LayoutInCell = 0;
            shape.RelativeHorizontalPosition = Word.WdRelativeHorizontalPosition.wdRelativeHorizontalPositionPage;
            shape.RelativeVerticalPosition = Word.WdRelativeVerticalPosition.wdRelativeVerticalPositionPage;
            shape.Left = left; shape.Top = y; shape.Width = w; shape.Height = 0;
            shape.Line.ForeColor.RGB = 0;
            shape.Line.Weight = 0.75f;
            shape.WrapFormat.Type = Word.WdWrapType.wdWrapFront;
            shape.LockAnchor = -1;
            return true;
        }
    }
}
