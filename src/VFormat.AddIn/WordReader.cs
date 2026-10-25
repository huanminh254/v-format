using System;
using System.Collections.Generic;
using VFormat.Core;
using Word = Microsoft.Office.Interop.Word;

namespace VFormat.AddIn
{
    /// <summary>Đọc tài liệu Word thành danh sách ParagraphInfo (kèm vị trí bảng, ô và định dạng hiện tại).</summary>
    public class WordSnapshot : IDisposable
    {
        public List<ParagraphInfo> Paragraphs = new List<ParagraphInfo>();
        public List<Word.Paragraph> WordParagraphs = new List<Word.Paragraph>();
        public List<int> Starts = new List<int>();

        public void Dispose()
        {
            if (WordParagraphs != null)
            {
                for (int i = 0; i < WordParagraphs.Count; i++)
                {
                    var p = WordParagraphs[i];
                    ComReleaser.Release(ref p);
                }
                WordParagraphs.Clear();
            }
        }
    }

    public static class WordReader
    {
        const int Undefined = 9999999;
        public const float PtPerMm = 72f / 25.4f;

        public static WordSnapshot Read(Word.Document doc, bool withFormat = true)
        {
            var snap = new WordSnapshot();

            // Bảng và ô cấp 1: (start, end, tableIndex, row, col)
            var cells = new List<(int s, int e, int t, int r, int c)>();
            int ti = 0;
            foreach (Word.Table table in doc.Tables)
            {
                ti++;
                try
                {
                    foreach (Word.Cell cell in table.Range.Cells)
                        if (cell.NestingLevel == 1)
                            cells.Add((cell.Range.Start, cell.Range.End, ti, cell.RowIndex, cell.ColumnIndex));
                }
                catch (Exception)
                {
                    cells.Add((table.Range.Start, table.Range.End, ti, 1, 1)); // bảng có ô gộp phức tạp
                }
            }

            cells.Sort((a, b) => a.s.CompareTo(b.s));
            int idx = 0, ci = 0;
            foreach (Word.Paragraph p in doc.Paragraphs)
            {
                var r = p.Range;
                var info = new ParagraphInfo { Index = idx++, Text = r.Text ?? "" };
                int start = r.Start;
                while (ci < cells.Count && cells[ci].e <= start) ci++;
                for (int j = ci; j < cells.Count && cells[j].s <= start; j++)
                    if (start >= cells[j].s && start < cells[j].e)
                    {
                        info.InTable = true; info.TableIndex = cells[j].t; info.CellRow = cells[j].r; info.CellColumn = cells[j].c;
                        break;
                    }

                if (TextUtil.Clean(info.Text).Length > 0)
                {
                    // thụt lề trái (khối chữ ký ở nửa phải) và số thứ tự tự động ("II." không có trong Range.Text)
                    var left = p.Format.LeftIndent;
                    info.LeftIndentPt = left >= Undefined ? 0 : left;
                    var lf = r.ListFormat;
                    if (lf.ListType != Word.WdListType.wdListNoNumbering)
                    {
                        info.InList = true;
                        var ls = lf.ListString;
                        info.ListPrefix = string.IsNullOrWhiteSpace(ls) ? null : ls;
                    }
                }
                if (withFormat && TextUtil.Clean(info.Text).Length > 0)
                {
                    var body = r.End - r.Start > 1 ? doc.Range(r.Start, r.End - 1) : r;
                    var f = body.Font;
                    var name = f.Name;
                    info.FontName = string.IsNullOrEmpty(name) ? "(nhiều phông chữ)" : name;
                    var size = f.Size;
                    info.FontSize = size >= Undefined ? 0 : size;
                    info.Bold = Tri(f.Bold);
                    info.Italic = Tri(f.Italic);
                    var color = (int)f.Color;
                    info.Colored = color != (int)Word.WdColor.wdColorAutomatic && color != 0;   // 0 = đen; hỗn hợp tính là có màu
                    var pf = p.Format;
                    info.FirstIndentPt = pf.FirstLineIndent >= Undefined ? 0 : pf.FirstLineIndent;
                    info.Alignment = ToAlign(pf.Alignment);
                }
                snap.Paragraphs.Add(info);
                snap.WordParagraphs.Add(p);
                snap.Starts.Add(start);
            }
            return snap;
        }

        static int Tri(int v) => v == Undefined ? -1 : v != 0 ? 1 : 0;

        static Align? ToAlign(Word.WdParagraphAlignment a)
        {
            switch (a)
            {
                case Word.WdParagraphAlignment.wdAlignParagraphLeft: return Align.Left;
                case Word.WdParagraphAlignment.wdAlignParagraphCenter: return Align.Center;
                case Word.WdParagraphAlignment.wdAlignParagraphRight: return Align.Right;
                case Word.WdParagraphAlignment.wdAlignParagraphJustify:
                case Word.WdParagraphAlignment.wdAlignParagraphDistribute:
                    return Align.Justify;
                default: return null;
            }
        }

        public static Word.WdParagraphAlignment ToWord(Align a)
        {
            switch (a)
            {
                case Align.Left: return Word.WdParagraphAlignment.wdAlignParagraphLeft;
                case Align.Center: return Word.WdParagraphAlignment.wdAlignParagraphCenter;
                case Align.Right: return Word.WdParagraphAlignment.wdAlignParagraphRight;
                default: return Word.WdParagraphAlignment.wdAlignParagraphJustify;
            }
        }

        public static PageInfo ReadPage(Word.Document doc)
        {
            var ps = doc.Sections[1].PageSetup;
            var info = new PageInfo
            {
                WidthMm = ps.PageWidth / PtPerMm, HeightMm = ps.PageHeight / PtPerMm,
                TopMm = ps.TopMargin / PtPerMm, BottomMm = ps.BottomMargin / PtPerMm,
                LeftMm = ps.LeftMargin / PtPerMm, RightMm = ps.RightMargin / PtPerMm,
                DifferentFirstPage = ps.DifferentFirstPageHeaderFooter != 0,
            };
            var s = doc.Sections[1];
            info.HasPageNumberInHeader = HasPageField(s.Headers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range);
            info.HasPageNumberInFooter = HasPageField(s.Footers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range);
            return info;
        }

        public static bool HasPageField(Word.Range r)
        {
            foreach (Word.Field f in r.Fields)
                if (f.Type == Word.WdFieldType.wdFieldPage) return true;
            // số trang dạng khung (PageNumbers.Add) nằm trong shape của header
            try
            {
                foreach (Word.Shape sh in r.ShapeRange)
                    if (sh.TextFrame.HasText != 0)
                        foreach (Word.Field f in sh.TextFrame.TextRange.Fields)
                            if (f.Type == Word.WdFieldType.wdFieldPage) return true;
            }
            catch { }
            return false;
        }
    }
}
