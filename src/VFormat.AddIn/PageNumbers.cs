using System;
using VFormat.Core;
using Word = Microsoft.Office.Interop.Word;

namespace VFormat.AddIn
{
    /// <summary>Số trang: chữ số Ả Rập, căn giữa đầu trang, không hiển thị ở trang thứ nhất (NĐ 30).</summary>
    public static class PageNumbers
    {
        public static void Apply(Word.Document doc, FormatProfile p)
        {
            bool first = true;
            foreach (Word.Section s in doc.Sections)
            {
                s.PageSetup.DifferentFirstPageHeaderFooter = first && p.HidePageNumberOnFirstPage ? -1 : 0;

                // Căn giữa số trang theo dải lề chuẩn kỹ thuật
                float sizePt = p.PageNumberSize > 0 ? p.PageNumberSize : 14f;
                if (p.PageNumberInHeader)
                {
                    float topMargin = s.PageSetup.TopMargin;
                    s.PageSetup.HeaderDistance = Math.Max(10f, (topMargin - sizePt) / 2f);
                }
                else
                {
                    float bottomMargin = s.PageSetup.BottomMargin;
                    s.PageSetup.FooterDistance = Math.Max(10f, (bottomMargin - sizePt) / 2f);
                }

                var primaryHeader = s.Headers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary];
                var primaryFooter = s.Footers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary];
                if (!first && primaryHeader.LinkToPrevious && primaryFooter.LinkToPrevious) continue;

                RemovePageFields(primaryHeader.Range);
                RemovePageFields(primaryFooter.Range);
                if (first)
                {
                    RemovePageFields(s.Headers[Word.WdHeaderFooterIndex.wdHeaderFooterFirstPage].Range);
                    RemovePageFields(s.Footers[Word.WdHeaderFooterIndex.wdHeaderFooterFirstPage].Range);
                }

                var target = (p.PageNumberInHeader ? primaryHeader : primaryFooter).Range;
                if (TextUtil.Clean(target.Text).Length == 0)
                {
                    target.Text = "";
                    target.Fields.Add(target, Word.WdFieldType.wdFieldPage);
                    target = (p.PageNumberInHeader ? primaryHeader : primaryFooter).Range;
                    target.Font.Name = p.FontName;
                    target.Font.Size = p.PageNumberSize;
                    target.Font.Bold = 0; target.Font.Italic = 0;
                    target.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                    target.ParagraphFormat.SpaceAfter = 0;
                    target.ParagraphFormat.SpaceBefore = 0;
                }
                s.Headers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].PageNumbers.RestartNumberingAtSection = false;
                first = false;
            }
        }

        static void RemovePageFields(Word.Range r)
        {
            for (int i = r.Fields.Count; i >= 1; i--)
            {
                var f = r.Fields[i];
                if (f.Type == Word.WdFieldType.wdFieldPage || f.Type == Word.WdFieldType.wdFieldNumPages) f.Delete();
            }
            for (int i = r.ShapeRange.Count; i >= 1; i--)
            {
                try
                {
                    var sh = r.ShapeRange[i];
                    if (sh.TextFrame.HasText != 0 && sh.TextFrame.TextRange.Fields.Count > 0
                        && sh.TextFrame.TextRange.Fields[1].Type == Word.WdFieldType.wdFieldPage) sh.Delete();
                }
                catch { }
            }
            // xóa đoạn chỉ còn "Trang" / "-" sau khi bỏ trường số trang
            var text = TextUtil.Clean(r.Text);
            if (text.Length > 0 && System.Text.RegularExpressions.Regex.IsMatch(text, @"^(Trang|Page|[\-–/ ]|của|of)*$"))
                r.Text = "";
        }
    }
}
