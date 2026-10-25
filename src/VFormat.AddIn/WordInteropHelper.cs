using System;
using System.Collections.Generic;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using Word = Microsoft.Office.Interop.Word;

namespace VFormat.AddIn
{
    /// <summary>
    /// Giải phóng đối tượng COM unmanaged an toàn, chống rò rỉ bộ nhớ và zombie WINWORD.EXE.
    /// </summary>
    public static class ComReleaser
    {
        public static void Release<T>(ref T comObj) where T : class
        {
            if (comObj == null) return;
            try
            {
                if (Marshal.IsComObject(comObj))
                {
                    Marshal.ReleaseComObject(comObj);
                }
            }
            catch { }
            finally
            {
                comObj = null;
            }
        }

        public static void FinalRelease<T>(ref T comObj) where T : class
        {
            if (comObj == null) return;
            try
            {
                if (Marshal.IsComObject(comObj))
                {
                    Marshal.FinalReleaseComObject(comObj);
                }
            }
            catch { }
            finally
            {
                comObj = null;
            }
        }

        public static void ReleaseAll(params object[] comObjs)
        {
            if (comObjs == null) return;
            foreach (var obj in comObjs)
            {
                if (obj == null) continue;
                try
                {
                    if (Marshal.IsComObject(obj))
                    {
                        Marshal.ReleaseComObject(obj);
                    }
                }
                catch { }
            }
        }
    }

    /// <summary>
    /// Phạm vi bảo tồn phông ký hiệu (Symbol, Wingdings, Webdings, Math).
    /// Ghi nhận các đoạn dùng font đặc biệt trước khi đổi font toàn bài và hoàn trả lại nguyên trạng sau khi gán.
    /// </summary>
    public sealed class SymbolFontScope : IDisposable
    {
        private readonly Word.Document _doc;
        private readonly List<(int Start, int End, string FontName)> _preservedSpans = new List<(int, int, string)>();

        private static readonly HashSet<string> SymbolFonts = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            "Symbol",
            "Wingdings",
            "Wingdings 2",
            "Wingdings 3",
            "Webdings",
            "MS Gothic",
            "MS Mincho",
            "Cambria Math",
            "Segoe UI Symbol",
            "Segoe UI Emoji",
            "Marlett"
        };

        public SymbolFontScope(Word.Document doc)
        {
            _doc = doc;
            if (_doc == null) return;

            try
            {
                ScanAndPreserve();
            }
            catch { }
        }

        private void ScanAndPreserve()
        {
            try
            {
                var content = _doc.Content;
                if (content == null || content.Characters.Count == 0) return;

                // Quét nhanh qua Find đối với từng font đặc biệt
                foreach (var fontName in SymbolFonts)
                {
                    var find = _doc.Content.Find;
                    find.ClearFormatting();
                    find.Font.Name = fontName;
                    find.Text = "";
                    find.Forward = true;
                    find.Wrap = Word.WdFindWrap.wdFindStop;
                    find.Format = true;

                    while (find.Execute())
                    {
                        var foundRange = find.Parent as Word.Range;
                        if (foundRange != null && foundRange.Start < foundRange.End)
                        {
                            _preservedSpans.Add((foundRange.Start, foundRange.End, fontName));
                        }
                    }
                }
            }
            catch { }
        }

        public void Dispose()
        {
            if (_doc == null || _preservedSpans.Count == 0) return;

            try
            {
                foreach (var span in _preservedSpans)
                {
                    try
                    {
                        if (span.Start < _doc.Content.End && span.End <= _doc.Content.End)
                        {
                            var r = _doc.Range(span.Start, span.End);
                            r.Font.Name = span.FontName;
                            r.Font.NameAscii = span.FontName;
                            r.Font.NameOther = span.FontName;
                        }
                    }
                    catch { }
                }
            }
            catch { }
        }
    }

    /// <summary>
    /// Kiểm tra trạng thái xem, Protected View và khoá văn bản trước khi can thiệp DOM.
    /// </summary>
    public static class WordProtectionHelper
    {
        public static void EnsureEditablePrintView(Word.Document doc)
        {
            if (doc == null) return;

            try
            {
                if (doc.ActiveWindow != null && doc.ActiveWindow.View != null)
                {
                    if (doc.ActiveWindow.View.Type != Word.WdViewType.wdPrintView)
                    {
                        doc.ActiveWindow.View.Type = Word.WdViewType.wdPrintView;
                    }
                }
            }
            catch { }

            try
            {
                if (doc.ProtectionType != Word.WdProtectionType.wdNoProtection)
                {
                    throw new InvalidOperationException(
                        "Tài liệu đang ở chế độ bảo vệ chống sửa đổi (Protected Document).\n" +
                        "Vui lòng vào tab Review -> Restrict Editing (Bảo vệ tài liệu) để tắt bảo vệ trước khi chuẩn hóa.");
                }
            }
            catch (COMException)
            {
                // Một số tài liệu IRM / Information Rights Management
                throw new InvalidOperationException(
                    "Tài liệu đang bị khoá quyền sửa đổi (Read-Only hoặc Protected View).\n" +
                    "Vui lòng bấm 'Enable Editing' trước khi thực hiện.");
            }
        }
    }

    /// <summary>
    /// Các tiện ích thao tác trực tiếp trên DOM của Word tương đương Tafinex WordToolEngine.
    /// </summary>
    public static class WordToolHelper
    {
        public static bool IsLikelyLegacyEncoded(string text)
        {
            if (string.IsNullOrEmpty(text)) return false;
            // Dấu hiệu đặc trưng TCVN3: có các ký tự đặc biệt đứng liền chữ cái hoặc các nguyên âm mã thấp
            return Regex.IsMatch(text, @"[\u0080-\u00FF\u0001-\u001F]") &&
                   (text.Contains("CéNG HßA") || text.Contains("VIÖT NAM") || text.Contains("®éc lËp") || text.Contains("quü") || text.Contains("tØnh"));
        }

        public static void CleanSpacingErrors(Word.Document doc, Word.Selection sel = null)
        {
            if (doc == null) return;
            var rng = (sel != null && sel.Type == Word.WdSelectionType.wdSelectionNormal && sel.Range.Characters.Count > 2)
                ? sel.Range
                : doc.Content;

            using (new WordBatch(doc.Application, "V-Format: Xóa khoảng trắng lỗi"))
            {
                // Bảo vệ dấu chấm lửng
                ReplaceText(rng, " ... ", " ___VF_BLANK___ ");

                // Bảo vệ khoảng trống ngày tháng
                for (int num = 15; num >= 2; num--)
                {
                    string sp = new string(' ', num);
                    ReplaceText(rng, "ngày" + sp + "tháng", "ngày___VF_SP1_" + num + "___tháng");
                    ReplaceText(rng, "tháng" + sp + "năm", "tháng___VF_SP2_" + num + "___năm");
                    ReplaceText(rng, "Số:" + sp + "/", "Số:___VF_NUMSP_" + num + "___/");
                    ReplaceText(rng, "Số: " + sp + "/", "Số: ___VF_NUMSP_" + num + "___/");
                }

                // Chuyển non-breaking space sang space thường
                ReplaceText(rng, "\u00a0", " ");

                // Dọn khoảng trắng thừa liên tiếp
                for (int i = 0; i < 4; i++)
                {
                    ReplaceText(rng, "  ", " ");
                }

                // Dọn khoảng trắng đứng trước dấu câu
                ReplaceText(rng, " ,", ",");
                ReplaceText(rng, " .", ".");
                ReplaceText(rng, " ;", ";");
                ReplaceText(rng, " :", ":");
                ReplaceText(rng, " ?", "?");
                ReplaceText(rng, " !", "!");
                ReplaceText(rng, " )", ")");
                ReplaceText(rng, " ]", "]");
                ReplaceText(rng, " }", "}");
                ReplaceText(rng, " ”", "”");
                ReplaceText(rng, " ’", "’");
                ReplaceText(rng, "( ", "(");
                ReplaceText(rng, "[ ", "[");
                ReplaceText(rng, "{ ", "{");
                ReplaceText(rng, "“ ", "“");
                ReplaceText(rng, "‘ ", "‘");

                // Phục hồi khoảng trống ngày tháng và chấm lửng
                for (int num = 15; num >= 2; num--)
                {
                    string sp = new string(' ', num);
                    ReplaceText(rng, "ngày___VF_SP1_" + num + "___tháng", "ngày" + sp + "tháng");
                    ReplaceText(rng, "tháng___VF_SP2_" + num + "___năm", "tháng" + sp + "năm");
                    ReplaceText(rng, "Số:___VF_NUMSP_" + num + "___/", "Số:" + sp + "/");
                    ReplaceText(rng, "Số: ___VF_NUMSP_" + num + "___/", "Số: " + sp + "/");
                }
                ReplaceText(rng, " ___VF_BLANK___ ", " ... ");
            }
        }

        public static int DeleteBlankPages(Word.Document doc)
        {
            if (doc == null) return 0;
            int count = 0;
            using (new WordBatch(doc.Application, "V-Format: Xóa trang trắng thừa"))
            {
                // Quét từ cuối văn bản ngược lên
                for (int i = doc.Paragraphs.Count; i >= 1; i--)
                {
                    try
                    {
                        var p = doc.Paragraphs[i];
                        string text = (p.Range.Text ?? "").Replace("\r", "").Replace("\a", "").Replace("\v", "").Trim();
                        if (string.IsNullOrEmpty(text))
                        {
                            // Kiểm tra không nằm trong bảng
                            if (!(bool)p.Range.Information[Word.WdInformation.wdWithInTable])
                            {
                                p.Range.Delete();
                                count++;
                            }
                        }
                        else
                        {
                            break; // Dừng khi gặp đoạn có nội dung
                        }
                    }
                    catch { }
                }
            }
            return count;
        }

        public static void ConvertTcvn3ToUnicode(Word.Document doc)
        {
            if (doc == null) return;
            using (new WordBatch(doc.Application, "V-Format: Chuyển mã TCVN3 sang Unicode"))
            {
                var map = new (string From, string To)[]
                {
                    ("CéNG HßA X· HéI CHñ NGHÜA VIÖT NAM", "CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM"),
                    ("Céng hßa x· héi chñ nghÜa ViÖt Nam", "Cộng hòa xã hội chủ nghĩa Việt Nam"),
                    ("§éc lËp - Tù do - H¹nh phóc", "Độc lập - Tự do - Hạnh phúc"),
                    ("§éC LËP - Tù DO - H¹NH PHóC", "ĐỘC LẬP - TỰ DO - HẠNH PHÚC"),
                    ("KÝnh göi:", "Kính gửi:"),
                    ("N¬i nhËn:", "Nơi nhận:"),
                    ("QUYÕT §ÞNH", "QUYẾT ĐỊNH"),
                    ("QuyÕt ®Þnh", "Quyết định"),
                    ("TH¤NG B ̧O", "THÔNG BÁO"),
                    ("Th«ng b ̧o", "Thông báo"),
                    ("B ̧O C ̧O", "BÁO CÁO"),
                    ("B ̧o c ̧o", "Báo cáo"),
                    ("Tê tr×nh", "Tờ trình"),
                    ("Tê TR×NH", "TỜ TRÌNH"),
                    ("C«ng v ̈n", "Công văn"),
                    ("Biªn b¶n", "Biên bản"),
                    ("Chñ tÞch", "Chủ tịch"),
                    ("Gi ̧m ®èc", "Giám đốc"),
                    ("Thñ tr−ëng", "Thủ trưởng"),
                    ("Tr−ëng phßng", "Trưởng phòng")
                };

                foreach (var pair in map)
                {
                    ReplaceText(doc.Content, pair.From, pair.To);
                }

                // Gán lại font Times New Roman
                try
                {
                    doc.Content.Font.Name = "Times New Roman";
                    doc.Content.Font.NameAscii = "Times New Roman";
                    doc.Content.Font.NameOther = "Times New Roman";
                }
                catch { }
            }
        }

        public static void ReplaceText(Word.Range rng, string findText, string replaceText)
        {
            if (rng == null || string.IsNullOrEmpty(findText)) return;
            try
            {
                var find = rng.Find;
                find.ClearFormatting();
                find.Replacement.ClearFormatting();
                find.Text = findText;
                find.Replacement.Text = replaceText;
                find.Forward = true;
                find.Wrap = Word.WdFindWrap.wdFindContinue;
                find.Format = false;
                find.MatchCase = true;
                find.MatchWholeWord = false;
                find.MatchWildcards = false;

                object replaceAll = Word.WdReplace.wdReplaceAll;
                find.Execute(Replace: ref replaceAll);
            }
            catch { }
        }
    }
}
