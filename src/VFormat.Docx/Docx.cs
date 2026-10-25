using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using VFormat.Core;

namespace VFormat.Docx
{
    /// <summary>Đọc/ghi .docx (OpenXML) tương đương WordReader + Formatter của add-in.</summary>
    public class DocxModel
    {
        public List<Paragraph> Paras = new List<Paragraph>();
        public List<ParagraphInfo> Infos = new List<ParagraphInfo>();
        public List<Table> Tables = new List<Table>();

        public static DocxModel Read(WordprocessingDocument doc)
        {
            var m = new DocxModel();
            var body = doc.MainDocumentPart.Document.Body;
            var defaults = Defaults.From(doc);
            m._numbering = new Numbering(doc);
            foreach (var el in body.Elements())
            {
                if (el is Paragraph p) m.Add(p, defaults, 0, 0, 0);
                else if (el is Table t)
                {
                    m.Tables.Add(t);
                    int ti = m.Tables.Count, r = 0;
                    foreach (var row in t.Elements<TableRow>())
                    {
                        r++; int c = 0;
                        foreach (var cell in row.Elements<TableCell>())
                        {
                            c++;
                            foreach (var cp in cell.Elements<Paragraph>()) m.Add(cp, defaults, ti, r, c);
                        }
                    }
                }
            }
            return m;
        }

        Numbering _numbering;

        void Add(Paragraph p, Defaults d, int table, int row, int col)
        {
            var info = new ParagraphInfo
            {
                Index = Paras.Count, Text = TextOf(p), InTable = table > 0, TableIndex = table, CellRow = row, CellColumn = col,
            };
            var numPr = p.ParagraphProperties?.NumberingProperties;
            if (numPr?.NumberingId?.Val?.Value is int numId && numId > 0)
            {
                info.InList = true;
                info.ListPrefix = _numbering.Next(numId, numPr.NumberingLevelReference?.Val?.Value ?? 0);
            }
            var pind = p.ParagraphProperties?.Indentation;
            var leftTw = pind?.Left?.Value ?? pind?.Start?.Value;
            if (leftTw != null && float.TryParse(leftTw, out var lt)) info.LeftIndentPt = lt / 20;
            var runs = p.Descendants<Run>().Where(r => r.Descendants<Text>().Any(t => t.Text.Trim().Length > 0)).ToList();
            if (runs.Count > 0)
            {
                var fonts = runs.Select(r => r.RunProperties?.RunFonts?.Ascii?.Value ?? d.Font).Distinct().ToList();
                info.FontName = fonts.Count == 1 ? fonts[0] : "(nhiều phông chữ)";
                var sizes = runs.Select(r => r.RunProperties?.FontSize?.Val?.Value is string s ? float.Parse(s) / 2 : d.Size).Distinct().ToList();
                info.FontSize = sizes.Count == 1 ? sizes[0] : 0;
                info.Bold = Tri(runs.Select(r => OnOff(r.RunProperties?.Bold)));
                info.Italic = Tri(runs.Select(r => OnOff(r.RunProperties?.Italic)));
                info.Colored = runs.Any(r => r.RunProperties?.Color is Color c &&
                    (c.Val?.Value is string v && v != "auto" && v != "000000" || c.ThemeColor != null && c.ThemeColor.Value != ThemeColorValues.Text1));
                var jc = p.ParagraphProperties?.Justification?.Val?.Value;
                info.Alignment = jc == JustificationValues.Center ? Align.Center
                    : jc == JustificationValues.Right || jc == JustificationValues.End ? Align.Right
                    : jc == JustificationValues.Both || jc == JustificationValues.Distribute ? Align.Justify : Align.Left;
                var ind = p.ParagraphProperties?.Indentation;
                if (ind?.FirstLine?.Value is string fl) info.FirstIndentPt = float.Parse(fl) / 20;
                else if (ind?.Hanging?.Value is string hg) info.FirstIndentPt = -float.Parse(hg) / 20;
            }
            Paras.Add(p); Infos.Add(info);
        }

        static bool OnOff(OnOffType t) => t != null && (t.Val == null || t.Val.Value);
        static int Tri(IEnumerable<bool> v) { var l = v.Distinct().ToList(); return l.Count != 1 ? -1 : l[0] ? 1 : 0; }

        public static string TextOf(Paragraph p)
        {
            var sb = new StringBuilder();
            foreach (var e in p.Descendants())
            {
                if (e is Text t) sb.Append(t.Text);
                else if (e is TabChar) sb.Append('\t');
                else if (e is Break) sb.Append('\v');
            }
            return sb.ToString();
        }

        /// <summary>Số thứ tự tự động của danh sách Word (numbering.xml): "II.", "1.", "a)"; dấu đầu dòng → "•".</summary>
        class Numbering
        {
            readonly Dictionary<int, AbstractNum> _abs = new Dictionary<int, AbstractNum>();
            readonly Dictionary<int, int> _numToAbs = new Dictionary<int, int>();
            readonly Dictionary<int, int[]> _counters = new Dictionary<int, int[]>();

            public Numbering(WordprocessingDocument doc)
            {
                var n = doc.MainDocumentPart.NumberingDefinitionsPart?.Numbering;
                if (n == null) return;
                foreach (var a in n.Elements<AbstractNum>())
                    if (a.AbstractNumberId?.Value is int id) _abs[id] = a;
                foreach (var ni in n.Elements<NumberingInstance>())
                    if (ni.NumberID?.Value is int id && ni.AbstractNumId?.Val?.Value is int aid) _numToAbs[id] = aid;
            }

            Level LevelOf(int absId, int lvl) =>
                _abs.TryGetValue(absId, out var a) ? a.Elements<Level>().FirstOrDefault(l => l.LevelIndex?.Value == lvl) : null;

            public string Next(int numId, int lvl)
            {
                if (!_numToAbs.TryGetValue(numId, out var absId) || lvl < 0 || lvl > 8) return null;
                var level = LevelOf(absId, lvl);
                if (level == null) return null;
                if (!_counters.TryGetValue(absId, out var c)) _counters[absId] = c = new int[9];
                c[lvl] = c[lvl] == 0 ? (level.StartNumberingValue?.Val?.Value ?? 1) : c[lvl] + 1;
                for (int i = lvl + 1; i < 9; i++) c[i] = 0;
                var fmt = level.NumberingFormat?.Val?.Value;
                if (fmt == NumberFormatValues.Bullet) return "•";
                var text = level.LevelText?.Val?.Value ?? "";
                return Regex.Replace(text, @"%(\d)", m =>
                {
                    int k = int.Parse(m.Groups[1].Value) - 1;
                    var lv = LevelOf(absId, k);
                    int v = c[k] == 0 ? (lv?.StartNumberingValue?.Val?.Value ?? 1) : c[k];
                    return Format(v, lv?.NumberingFormat?.Val?.Value);
                });
            }

            static string Format(int v, NumberFormatValues? f)
            {
                if (f == NumberFormatValues.UpperRoman) return Roman(v);
                if (f == NumberFormatValues.LowerRoman) return Roman(v).ToLowerInvariant();
                if (f == NumberFormatValues.UpperLetter) return ((char)('A' + (v - 1) % 26)).ToString();
                if (f == NumberFormatValues.LowerLetter) return ((char)('a' + (v - 1) % 26)).ToString();
                return v.ToString();
            }

            static string Roman(int v)
            {
                var sb = new StringBuilder();
                foreach (var (n, r) in new[] { (1000, "M"), (900, "CM"), (500, "D"), (400, "CD"), (100, "C"), (90, "XC"),
                                               (50, "L"), (40, "XL"), (10, "X"), (9, "IX"), (5, "V"), (4, "IV"), (1, "I") })
                    while (v >= n) { sb.Append(r); v -= n; }
                return sb.ToString();
            }
        }

        class Defaults
        {
            public string Font = "Calibri"; public float Size = 11;
            public static Defaults From(WordprocessingDocument doc)
            {
                var d = new Defaults();
                var styles = doc.MainDocumentPart.StyleDefinitionsPart?.Styles;
                var theme = doc.MainDocumentPart.ThemePart?.Theme;
                string minor = theme?.ThemeElements?.FontScheme?.MinorFont?.LatinFont?.Typeface?.Value;
                void FromRpr(OpenXmlElement rpr)
                {
                    if (rpr == null) return;
                    var f = rpr.GetFirstChild<RunFonts>();
                    if (f?.Ascii?.Value != null) d.Font = f.Ascii.Value;
                    else if (f?.AsciiTheme != null && minor != null) d.Font = minor;
                    var sz = rpr.GetFirstChild<FontSize>();
                    if (sz?.Val?.Value != null) d.Size = float.Parse(sz.Val.Value) / 2;
                }
                FromRpr(styles?.DocDefaults?.RunPropertiesDefault?.RunPropertiesBaseStyle);
                FromRpr(styles?.Elements<Style>().FirstOrDefault(s => s.Default?.Value == true && s.Type?.Value == StyleValues.Paragraph)?.StyleRunProperties);
                return d;
            }
        }
    }

    /// <summary>Dọn văn bản chuyển từ PDF / gõ Shift+Enter (tương đương Formatter.CleanupConvertedText của add-in).</summary>
    public static class DocxCleanup
    {
        static bool IsLineBreak(Break b) => b.Type == null || b.Type.Value == BreakValues.TextWrapping;

        public static bool Run(DocxModel m)
        {
            bool converted = ConvertedText.LooksConverted(m.Infos);
            bool changed = false;
            int n = m.Paras.Count;
            var pageNumbers = new HashSet<int>();
            if (converted)
                for (int i = 0; i < n; i++)
                {
                    if (m.Infos[i].InTable || m.Paras[i].ParagraphProperties?.SectionProperties == null) continue;
                    int j = i + 1;
                    while (j < n && !m.Infos[j].InTable && TextUtil.Clean(m.Infos[j].Text).Length == 0) j++;
                    if (j < n && !m.Infos[j].InTable && ConvertedText.IsPageNumberArtifact(m.Infos[j].Text)) pageNumbers.Add(j);
                }

            for (int i = n - 1; i >= 0; i--)
            {
                if (m.Infos[i].InTable) continue;
                var p = m.Paras[i];
                if (p.Descendants<FieldCode>().Any() || p.Descendants<SimpleField>().Any() || p.Descendants<Drawing>().Any()) continue;
                if (pageNumbers.Contains(i) && p.ParagraphProperties?.SectionProperties == null) { p.Remove(); changed = true; continue; }
                var header = i < 20 ? ConvertedText.SplitMergedHeader(m.Infos[i].Text) : null;
                if (header != null)
                {
                    var prev = p;
                    SetTextWithTabs(p, header[0]);
                    for (int k = 1; k < header.Length; k++)
                    {
                        var np = new Paragraph();
                        if (p.ParagraphProperties != null)
                        {
                            var ppr = (ParagraphProperties)p.ParagraphProperties.CloneNode(true);
                            ppr.SectionProperties = null;
                            np.AppendChild(ppr);
                        }
                        SetTextWithTabs(np, header[k]);
                        prev.InsertAfterSelf(np); prev = np;
                    }
                    changed = true;
                    continue;
                }
                changed |= Reflow(p);
                // tab / khoảng trắng đầu đoạn dựng vị trí chữ của bản PDF
                if (converted) changed |= TrimLeading(p);
            }

            if (converted)
            {
                var sects = m.Paras.Select(x => x.ParagraphProperties?.SectionProperties).Where(x => x != null).ToList();
                var last = m.Paras.FirstOrDefault()?.Ancestors<Body>().FirstOrDefault()?.Elements<SectionProperties>().LastOrDefault();
                if (last != null) sects.Add(last);
                // các section sau section đầu: liền mạch, không sang trang
                foreach (var sp in sects.Skip(1))
                {
                    sp.RemoveAllChildren<SectionType>();
                    var type = new SectionType { Val = SectionMarkValues.Continuous };
                    var size = sp.GetFirstChild<PageSize>();
                    if (size != null) size.InsertBeforeSelf(type); else sp.AppendChild(type);
                }
            }
            return changed;
        }

        /// <summary>Nối / tách các ngắt dòng cứng trong đoạn theo ConvertedText.Plan.</summary>
        static bool Reflow(Paragraph p)
        {
            var breaks = p.Descendants<Break>().Where(IsLineBreak).ToList();
            if (breaks.Count == 0) return false;
            var sb = new StringBuilder();
            foreach (var e in p.Descendants())
            {
                if (e is Text t) sb.Append(t.Text);
                else if (e is TabChar) sb.Append('\t');
                else if (e is Break b && IsLineBreak(b)) sb.Append(ConvertedText.LineBreak);
            }
            var raw = sb.ToString();
            var plan = ConvertedText.Plan(raw);
            var lines = raw.Split(ConvertedText.LineBreak);
            for (int k = breaks.Count - 1; k >= 0; k--)
            {
                var br = breaks[k];
                bool spaced = lines[k].EndsWith(" ") || lines[k + 1].StartsWith(" ");
                var action = plan[k];
                var run = br.Parent as Run;
                if (action == LineBreakAction.Split && (run == null || run.Parent != p)) action = LineBreakAction.Join;
                if (action == LineBreakAction.Split) SplitAt(p, run, br);
                else
                {
                    if (action == LineBreakAction.Join && !spaced)
                        br.InsertAfterSelf(new Text(" ") { Space = SpaceProcessingModeValues.Preserve });
                    br.Remove();
                }
            }
            return true;
        }

        /// <summary>Tách đoạn p tại ngắt dòng br (nằm trong run con trực tiếp của p): phần sau thành đoạn mới ngay sau p.</summary>
        static void SplitAt(Paragraph p, Run run, Break br)
        {
            var tail = new Run();
            if (run.RunProperties != null) tail.AppendChild(run.RunProperties.CloneNode(true));
            foreach (var e in br.ElementsAfter().ToList()) { e.Remove(); tail.AppendChild(e); }
            br.Remove();
            var np = new Paragraph();
            if (p.ParagraphProperties != null)
            {
                np.AppendChild(p.ParagraphProperties.CloneNode(true));
                p.ParagraphProperties.SectionProperties = null;   // dấu ngắt section ở cuối đoạn: đi theo phần sau
            }
            np.AppendChild(tail);
            foreach (var e in run.ElementsAfter().ToList()) { e.Remove(); np.AppendChild(e); }
            p.InsertAfterSelf(np);
            TrimLeading(np);
        }

        /// <summary>Bỏ tab / khoảng trắng / ngắt dòng ở đầu đoạn (trước chữ đầu tiên). Không đụng đoạn chỉ có khoảng trắng.</summary>
        static bool TrimLeading(Paragraph p)
        {
            if (TextUtil.Clean(DocxModel.TextOf(p)).Length == 0) return false;
            bool changed = false;
            foreach (var e in p.Descendants().ToList())
            {
                if (e is TabChar || e is Break b && IsLineBreak(b)) { e.Remove(); changed = true; continue; }
                if (e is Text t)
                {
                    var s = t.Text.TrimStart(' ', '\t');
                    if (s.Length != t.Text.Length) { t.Text = s; t.Space = SpaceProcessingModeValues.Preserve; changed = true; }
                    if (s.Length > 0) break;
                }
            }
            return changed;
        }

        /// <summary>Thay chữ của đoạn bằng text (ký tự '\t' thành tab thật), bỏ chỉ số trên/dưới của run đầu.</summary>
        static void SetTextWithTabs(Paragraph p, string text)
        {
            var firstRun = p.Descendants<Run>().FirstOrDefault(r => r.Descendants<Text>().Any());
            var rpr = (RunProperties)firstRun?.RunProperties?.CloneNode(true);
            if (rpr != null) rpr.VerticalTextAlignment = null;
            foreach (var e in p.Elements().Where(e => !(e is ParagraphProperties)).ToList()) e.Remove();
            var parts = text.Split('\t');
            for (int k = 0; k < parts.Length; k++)
            {
                var run = new Run();
                if (rpr != null) run.AppendChild(rpr.CloneNode(true));
                if (k > 0) run.AppendChild(new TabChar());
                run.AppendChild(new Text(parts[k]) { Space = SpaceProcessingModeValues.Preserve });
                p.AppendChild(run);
            }
        }
    }

    public static class DocxWriter
    {
        public const int TwipPerPt = 20;
        public static int MmToTwip(float mm) => (int)Math.Round(mm * 1440 / 25.4);

        // ---------- trang ----------

        /// <param name="landscapeAllowed">false: trang ngang được xoay về dọc (văn bản không có bảng nội dung)</param>
        public static void PageSetup(WordprocessingDocument doc, FormatProfile p, bool landscapeAllowed = true)
        {
            foreach (var sp in doc.MainDocumentPart.Document.Body.Descendants<SectionProperties>().ToList())
            {
                var size = sp.GetFirstChild<PageSize>() ?? sp.AppendChild(new PageSize());
                bool landscape = size.Orient?.Value == PageOrientationValues.Landscape
                    || size.Orient == null && size.Width?.Value > size.Height?.Value;
                if (landscape && !landscapeAllowed) { size.Orient = null; landscape = false; }
                else if (landscape) size.Orient = PageOrientationValues.Landscape;
                size.Width = (UInt32Value)(uint)MmToTwip(landscape ? p.PageHeightMm : p.PageWidthMm);
                size.Height = (UInt32Value)(uint)MmToTwip(landscape ? p.PageWidthMm : p.PageHeightMm);
                var mg = sp.GetFirstChild<PageMargin>() ?? sp.AppendChild(new PageMargin());
                mg.Top = MmToTwip(p.MarginTopMm); mg.Bottom = MmToTwip(p.MarginBottomMm);
                mg.Left = (uint)MmToTwip(p.MarginLeftMm); mg.Right = (uint)MmToTwip(p.MarginRightMm);
                mg.Header = (uint)MmToTwip(10); mg.Footer = (uint)MmToTwip(10); mg.Gutter = 0;
            }
        }

        public static PageInfo ReadPage(WordprocessingDocument doc)
        {
            var sp = doc.MainDocumentPart.Document.Body.Elements<SectionProperties>().LastOrDefault();
            var size = sp?.GetFirstChild<PageSize>(); var mg = sp?.GetFirstChild<PageMargin>();
            float Mm(double tw) => (float)(tw * 25.4 / 1440);
            var info = new PageInfo
            {
                WidthMm = Mm(size?.Width?.Value ?? 12240), HeightMm = Mm(size?.Height?.Value ?? 15840),
                TopMm = Mm(mg?.Top?.Value ?? 1440), BottomMm = Mm(mg?.Bottom?.Value ?? 1440),
                LeftMm = Mm(mg?.Left?.Value ?? 1440), RightMm = Mm(mg?.Right?.Value ?? 1440),
                DifferentFirstPage = sp?.GetFirstChild<TitlePage>() != null,
            };
            bool HasPage(OpenXmlPartRootElement root) => root != null &&
                (root.Descendants<SimpleField>().Any(f => f.Instruction?.Value?.Contains("PAGE") == true) ||
                 root.Descendants<FieldCode>().Any(f => f.Text.Contains("PAGE")));
            foreach (var hr in sp?.Elements<HeaderReference>() ?? Enumerable.Empty<HeaderReference>())
                if (hr.Type?.Value != HeaderFooterValues.First && HasPage(((HeaderPart)doc.MainDocumentPart.GetPartById(hr.Id)).Header))
                    info.HasPageNumberInHeader = true;
            foreach (var fr in sp?.Elements<FooterReference>() ?? Enumerable.Empty<FooterReference>())
                if (fr.Type?.Value != HeaderFooterValues.First && HasPage(((FooterPart)doc.MainDocumentPart.GetPartById(fr.Id)).Footer))
                    info.HasPageNumberInFooter = true;
            return info;
        }

        /// <summary>Số trang giữa đầu trang (hoặc chân trang với KHBD), ẩn ở trang đầu.</summary>
        public static void PageNumbers(WordprocessingDocument doc, FormatProfile p)
        {
            var main = doc.MainDocumentPart;
            var sp = main.Document.Body.Elements<SectionProperties>().Last();
            Paragraph NumberPara() => new Paragraph(
                new ParagraphProperties
                {
                    SpacingBetweenLines = new SpacingBetweenLines { Before = "0", After = "0" },
                    Justification = new Justification { Val = JustificationValues.Center },
                },
                new SimpleField(new Run(RunProps(p.FontName, p.PageNumberSize), new Text("1"))) { Instruction = " PAGE " });

            foreach (var r in sp.Elements<HeaderReference>().ToList()) { main.DeletePart(r.Id); r.Remove(); }
            foreach (var r in sp.Elements<FooterReference>().ToList()) { main.DeletePart(r.Id); r.Remove(); }

            OpenXmlElement firstRef, defRef;
            if (p.PageNumberInHeader)
            {
                var hp = main.AddNewPart<HeaderPart>(); hp.Header = new Header(NumberPara());
                var hf = main.AddNewPart<HeaderPart>(); hf.Header = new Header(new Paragraph());
                defRef = new HeaderReference { Type = HeaderFooterValues.Default, Id = main.GetIdOfPart(hp) };
                firstRef = new HeaderReference { Type = HeaderFooterValues.First, Id = main.GetIdOfPart(hf) };
            }
            else
            {
                var fp = main.AddNewPart<FooterPart>(); fp.Footer = new Footer(NumberPara());
                var ff = main.AddNewPart<FooterPart>(); ff.Footer = new Footer(new Paragraph());
                defRef = new FooterReference { Type = HeaderFooterValues.Default, Id = main.GetIdOfPart(fp) };
                firstRef = new FooterReference { Type = HeaderFooterValues.First, Id = main.GetIdOfPart(ff) };
            }
            sp.InsertAt(firstRef, 0);
            sp.InsertAt(defRef, 0);
            if (p.HidePageNumberOnFirstPage && sp.GetFirstChild<TitlePage>() == null)
            {
                // titlePg phải đứng sau pgNumType/cols/... — đặt trước textDirection/docGrid nếu có
                var before = (OpenXmlElement)sp.GetFirstChild<DocGrid>() ?? sp.GetFirstChild<TextDirection>();
                if (before != null) sp.InsertBefore(new TitlePage(), before); else sp.AppendChild(new TitlePage());
            }
        }

        public static RunProperties RunProps(string font, float size) => new RunProperties(
            new RunFonts { Ascii = font, HighAnsi = font, ComplexScript = font, EastAsia = font },
            new FontSize { Val = ((int)(size * 2)).ToString() }, new FontSizeComplexScript { Val = ((int)(size * 2)).ToString() });

        // ---------- bảng khối đầu / chữ ký ----------

        public static Table ReplaceWithTable(DocxModel m, TablePlan plan)
        {
            var first = m.Paras[plan.First];
            var table = NewTwoColumnTable(plan.Left, plan.Right, plan.LeftMm, plan.RightMm);
            first.InsertBeforeSelf(table);
            for (int i = plan.First; i <= plan.Last; i++) m.Paras[i].Remove();
            foreach (var i in plan.Remove) m.Paras[i].Remove();
            return table;
        }

        public static Table NewTwoColumnTable(IList<string> left, IList<string> right, float leftMm, float rightMm)
        {
            int lw = MmToTwip(leftMm), rw = MmToTwip(rightMm);
            // dùng setter có kiểu để SDK tự xếp đúng thứ tự phần tử theo schema
            var tp = new TableProperties
            {
                TableWidth = new TableWidth { Width = (lw + rw).ToString(), Type = TableWidthUnitValues.Dxa },
                TableIndentation = new TableIndentation { Width = 0, Type = TableWidthUnitValues.Dxa },
                TableBorders = NoBorders(),
                TableLayout = new TableLayout { Type = TableLayoutValues.Fixed },
                TableCellMarginDefault = new TableCellMarginDefault(
                    new TopMargin { Width = "0", Type = TableWidthUnitValues.Dxa },
                    new TableCellLeftMargin { Width = 0, Type = TableWidthValues.Dxa },
                    new BottomMargin { Width = "0", Type = TableWidthUnitValues.Dxa },
                    new TableCellRightMargin { Width = 0, Type = TableWidthValues.Dxa }),
            };
            var t = new Table(tp,
                new TableGrid(new GridColumn { Width = lw.ToString() }, new GridColumn { Width = rw.ToString() }),
                new TableRow(new TableRowProperties(new CantSplit()), Cell(left, lw), Cell(right, rw)));
            return t;
        }

        static TableCell Cell(IList<string> lines, int width)
        {
            var c = new TableCell(new TableCellProperties(new TableCellWidth { Width = width.ToString(), Type = TableWidthUnitValues.Dxa }));
            foreach (var l in lines.DefaultIfEmpty(""))
                c.AppendChild(new Paragraph(new ParagraphProperties(new Justification { Val = JustificationValues.Center }),
                    new Run(new Text(l) { Space = SpaceProcessingModeValues.Preserve })));
            return c;
        }

        static TableBorders NoBorders() => new TableBorders(
            new TopBorder { Val = BorderValues.None }, new LeftBorder { Val = BorderValues.None },
            new BottomBorder { Val = BorderValues.None }, new RightBorder { Val = BorderValues.None },
            new InsideHorizontalBorder { Val = BorderValues.None }, new InsideVerticalBorder { Val = BorderValues.None });

        public static int FormatTables(DocxModel m, ClassifiedDocument cd, FormatProfile p, float availableMm)
        {
            var (header, sign) = Planner.FormalTables(cd);
            for (int i = 0; i < m.Tables.Count; i++)
            {
                var t = m.Tables[i];
                var tp = t.GetFirstChild<TableProperties>() ?? t.PrependChild(new TableProperties());
                if (header.Contains(i + 1) || sign.Contains(i + 1))
                {
                    tp.TableBorders = NoBorders();
                    var grid = t.GetFirstChild<TableGrid>()?.Elements<GridColumn>().ToList();
                    if (grid != null && grid.Count == 2)
                    {
                        float l = header.Contains(i + 1) ? p.HeaderLeftColMm : p.SignLeftColMm;
                        float r = header.Contains(i + 1) ? p.HeaderRightColMm : p.SignRightColMm;
                        grid[0].Width = MmToTwip(l).ToString(); grid[1].Width = MmToTwip(r).ToString();
                        foreach (var row in t.Elements<TableRow>())
                        {
                            var cells = row.Elements<TableCell>().ToList();
                            if (cells.Count != 2) continue;
                            SetCellWidth(cells[0], MmToTwip(l)); SetCellWidth(cells[1], MmToTwip(r));
                        }
                        tp.TableWidth = new TableWidth { Width = (MmToTwip(l) + MmToTwip(r)).ToString(), Type = TableWidthUnitValues.Dxa };
                    }
                    continue;
                }
                tp.TableWidth = new TableWidth { Width = "5000", Type = TableWidthUnitValues.Pct };
                tp.TableJustification = new TableJustification { Val = TableRowAlignmentValues.Center };
                tp.TableIndentation = null;
                tp.TableLayout = null;
                var widths = Planner.PlanColumnWidths(cd, i + 1, availableMm);
                var gridCols = t.GetFirstChild<TableGrid>()?.Elements<GridColumn>().ToList();
                if (widths != null && gridCols != null && gridCols.Count == widths.Length)
                {
                    // độ rộng cột theo nội dung (dùng chung với add-in), đủ khổ vùng chữ
                    for (int c = 0; c < widths.Length; c++) gridCols[c].Width = MmToTwip(widths[c]).ToString();
                    foreach (var row in t.Elements<TableRow>())
                    {
                        var cells = row.Elements<TableCell>().ToList();
                        for (int c = 0; c < cells.Count && c < widths.Length; c++) SetCellWidth(cells[c], MmToTwip(widths[c]));
                    }
                    tp.TableWidth = new TableWidth { Width = widths.Sum(w => MmToTwip(w)).ToString(), Type = TableWidthUnitValues.Dxa };
                    tp.TableLayout = new TableLayout { Type = TableLayoutValues.Fixed };
                }
                else
                {
                    // ô gộp: bỏ độ rộng cố định để bảng co vừa 100%
                    foreach (var gc in gridCols ?? Enumerable.Empty<GridColumn>().ToList()) gc.Width = null;
                    foreach (var tcw in t.Descendants<TableCellWidth>().ToList()) tcw.Remove();
                }
                var rows = t.Elements<TableRow>().ToList();
                if (rows.Count > 2)
                {
                    var trPr = rows[0].TableRowProperties ?? rows[0].PrependChild(new TableRowProperties());
                    if (trPr.GetFirstChild<TableHeader>() == null) trPr.AppendChild(new TableHeader());
                }
            }
            return m.Tables.Count;
        }

        static void SetCellWidth(TableCell c, int tw)
        {
            var pr = c.TableCellProperties ?? c.PrependChild(new TableCellProperties());
            pr.TableCellWidth = new TableCellWidth { Width = tw.ToString(), Type = TableWidthUnitValues.Dxa };
        }

        // ---------- chữ ----------

        public static void SetText(Paragraph p, string text)
        {
            var firstRun = p.Descendants<Run>().FirstOrDefault(r => r.Descendants<Text>().Any());
            var rpr = firstRun?.RunProperties?.CloneNode(true);
            foreach (var e in p.Elements().Where(e => !(e is ParagraphProperties)).ToList()) e.Remove();
            var run = new Run(new Text(text) { Space = SpaceProcessingModeValues.Preserve });
            if (rpr != null) run.PrependChild(rpr);
            p.AppendChild(run);
        }

        /// <summary>
        /// Sửa chính tả ngay trong các run của đoạn: chỉ thay đúng chuỗi lỗi, giữ định dạng ký tự (đậm, nghiêng...),
        /// tab, ngắt dòng, trường. Lỗi nằm vắt qua nhiều run thì gợi ý lấy định dạng của run đầu; lỗi vắt qua tab/ngắt
        /// dòng thì bỏ qua. Trả về số chỗ đã sửa.
        /// </summary>
        public static int FixSpelling(Paragraph p, SpellChecker sc)
        {
            var segs = new List<(Text node, int start, int length)>();
            var sb = new StringBuilder();
            foreach (var e in p.Descendants())
            {
                if (e is Text t)
                {
                    if (!t.Text.IsNormalized(NormalizationForm.FormC)) t.Text = t.Text.Normalize(NormalizationForm.FormC);
                    segs.Add((t, sb.Length, t.Text.Length));
                    sb.Append(t.Text);
                }
                else if (e is TabChar) sb.Append('\t');
                else if (e is Break) sb.Append('\n');
            }
            var text = sb.ToString();
            if (text.Trim().Length == 0) return 0;

            int n = 0;
            // từ cuối lên: phần đầu của mỗi node (trước chỗ đã sửa) giữ nguyên vị trí
            foreach (var h in sc.Check(text).OrderByDescending(h => h.Start))
            {
                int end = h.Start + h.Length;
                var covered = segs.Where(s => s.start < end && h.Start < s.start + s.length).ToList();
                int inText = covered.Sum(s => Math.Min(end, s.start + s.length) - Math.Max(h.Start, s.start));
                if (covered.Count == 0 || inText != h.Length) continue;

                // gợi ý dài bằng chuỗi lỗi (đề suất → đề xuất): mỗi run giữ đúng phần ký tự của nó, kể cả định dạng;
                // khác độ dài: cả gợi ý vào run đầu, các run sau bỏ phần thuộc lỗi
                bool sameLength = h.Suggestion.Length == h.Length;
                for (int k = 0; k < covered.Count; k++)
                {
                    var s = covered[k];
                    int from = Math.Max(h.Start, s.start) - s.start;
                    int to = Math.Min(end, s.start + s.length) - s.start;
                    string put = sameLength ? h.Suggestion.Substring(s.start + from - h.Start, to - from)
                                            : k == 0 ? h.Suggestion : "";
                    s.node.Text = s.node.Text.Substring(0, from) + put + s.node.Text.Substring(to);
                    s.node.Space = SpaceProcessingModeValues.Preserve;
                }
                n++;
            }
            return n;
        }

        /// <summary>Chèn đoạn mới chứa <paramref name="text"/> ngay sau đoạn p (giữ định dạng đoạn của p).</summary>
        public static Paragraph InsertParagraphAfter(Paragraph p, string text)
        {
            var np = new Paragraph();
            if (p.ParagraphProperties != null) np.AppendChild(p.ParagraphProperties.CloneNode(true));
            np.AppendChild(new Run(new Text(text) { Space = SpaceProcessingModeValues.Preserve }));
            p.InsertAfterSelf(np);
            return np;
        }

        static readonly Regex[] SpacingFixes =
        {
            new Regex(" {2,}"),
            new Regex(@" +(?=[,;:\?!\)]|\.(?!\.))"),
            new Regex(@"(?<=\() +"),
            new Regex(@"(?<=\d) +(?=/ *[A-ZĐ\d])"),   // "số 15 /KH-VP"
        };

        public static int FixSpacing(Paragraph p)
        {
            int n = 0;
            foreach (var t in p.Descendants<Text>())
            {
                var s = t.Text;
                s = SpacingFixes[0].Replace(s, " ");
                s = SpacingFixes[1].Replace(s, "");
                s = SpacingFixes[2].Replace(s, "");
                s = SpacingFixes[3].Replace(s, "");
                if (s != t.Text) { t.Text = s; t.Space = SpaceProcessingModeValues.Preserve; n++; }
            }
            return n;
        }

        // ---------- định dạng đoạn ----------

        public static void Apply(Paragraph p, ParagraphPlan plan, string font)
        {
            if (plan.BoldPrefixLength > 0) SplitAt(p, plan.BoldPrefixLength);
            if (plan.UnderlinePrefixLength > 0) SplitAt(p, plan.UnderlinePrefixLength);
            if (plan.ItalicPrefixLength > 0) SplitAt(p, plan.ItalicPrefixLength);

            var ppr = p.ParagraphProperties ?? p.PrependChild(new ParagraphProperties());
            if (plan.Align.HasValue)
            {
                ppr.Justification = new Justification
                {
                    Val = plan.Align == Align.Center ? JustificationValues.Center : plan.Align == Align.Right ? JustificationValues.Right
                        : plan.Align == Align.Left ? JustificationValues.Left : JustificationValues.Both,
                };
            }
            if (plan.FirstIndentPt.HasValue || plan.ResetSideIndents || plan.LeftIndentPt.HasValue)
            {
                var ind = ppr.Indentation ?? (ppr.Indentation = new Indentation());
                if (plan.ResetSideIndents) { ind.Left = "0"; ind.Right = "0"; ind.Start = null; ind.End = null; }
                if (plan.LeftIndentPt.HasValue)
                {
                    ind.Left = ((int)Math.Round(plan.LeftIndentPt.Value * TwipPerPt)).ToString();
                    ind.Right = "0"; ind.Start = null; ind.End = null;
                }
                if (plan.FirstIndentPt.HasValue) { ind.Hanging = null; ind.FirstLine = ((int)Math.Round(plan.FirstIndentPt.Value * TwipPerPt)).ToString(); }
            }
            ppr.SpacingBetweenLines = new SpacingBetweenLines
            {
                Before = ((int)(plan.SpaceBeforePt * TwipPerPt)).ToString(), After = ((int)(plan.SpaceAfterPt * TwipPerPt)).ToString(),
                Line = plan.LineExactPt.HasValue ? ((int)Math.Round(plan.LineExactPt.Value * TwipPerPt)).ToString()
                    : ((int)Math.Round(240 * (plan.LineMultiple ?? 1f))).ToString(),
                LineRule = plan.LineExactPt.HasValue ? LineSpacingRuleValues.Exact : LineSpacingRuleValues.Auto,
            };
            ppr.KeepNext = plan.KeepWithNext ? new KeepNext() : null;
            if (plan.BorderBottom)
                ppr.ParagraphBorders = new ParagraphBorders(new BottomBorder { Val = BorderValues.Single, Size = 6, Space = 4, Color = "000000" });

            // dấu đoạn (giữ chiều cao dòng trống đúng cỡ)
            var mark = ppr.ParagraphMarkRunProperties ?? (ppr.ParagraphMarkRunProperties = new ParagraphMarkRunProperties());
            mark.RemoveAllChildren<RunFonts>(); mark.RemoveAllChildren<FontSize>(); mark.RemoveAllChildren<FontSizeComplexScript>();
            mark.Append(new RunFonts { Ascii = font, HighAnsi = font, ComplexScript = font, EastAsia = font });
            if (plan.Size.HasValue) mark.Append(new FontSize { Val = ((int)(plan.Size * 2)).ToString() });

            int pos = 0;
            foreach (var run in p.Descendants<Run>().ToList())
            {
                var len = run.Descendants<Text>().Sum(t => t.Text.Length);
                var rpr = run.RunProperties ?? run.PrependChild(new RunProperties());
                rpr.RunFonts = new RunFonts { Ascii = font, HighAnsi = font, ComplexScript = font, EastAsia = font };
                if (plan.ResetColor) rpr.Color = null;
                if (plan.ClearUnderline) rpr.Underline = null;
                if (plan.Size.HasValue)
                {
                    rpr.FontSize = new FontSize { Val = ((int)(plan.Size * 2)).ToString() };
                    rpr.FontSizeComplexScript = new FontSizeComplexScript { Val = ((int)(plan.Size * 2)).ToString() };
                }
                bool? bold = plan.Bold;
                if (plan.BoldPrefixLength > 0) bold = pos < plan.BoldPrefixLength;
                if (bold.HasValue) rpr.Bold = bold.Value ? new Bold() : new Bold { Val = false };
                bool? italic = plan.Italic;
                if (plan.ItalicPrefixLength > 0) italic = pos < plan.ItalicPrefixLength;
                if (italic.HasValue) rpr.Italic = italic.Value ? new Italic() : new Italic { Val = false };
                if (plan.Align.HasValue) rpr.Underline = null;
                if (plan.UnderlinePrefixLength > 0 && pos < plan.UnderlinePrefixLength)
                    rpr.Underline = new Underline { Val = UnderlineValues.Single };
                if (plan.ToUpper)
                    foreach (var t in run.Descendants<Text>()) t.Text = TextUtil.Upper(t.Text);
                pos += len;
            }
        }

        /// <summary>Tách run sao cho có ranh giới đúng tại vị trí offset (tính theo ký tự chữ).</summary>
        static void SplitAt(Paragraph p, int offset)
        {
            int pos = 0;
            foreach (var run in p.Descendants<Run>().ToList())
            {
                var texts = run.Elements<Text>().ToList();
                int len = texts.Sum(t => t.Text.Length);
                if (offset > pos && offset < pos + len && texts.Count == 1)
                {
                    var t = texts[0]; int cut = offset - pos;
                    var second = (Run)run.CloneNode(true);
                    second.GetFirstChild<Text>().Text = t.Text.Substring(cut);
                    second.GetFirstChild<Text>().Space = SpaceProcessingModeValues.Preserve;
                    t.Text = t.Text.Substring(0, cut); t.Space = SpaceProcessingModeValues.Preserve;
                    run.InsertAfterSelf(second);
                    return;
                }
                pos += len;
            }
        }

        // ---------- đường kẻ ----------

        /// <summary>
        /// Kẻ đường bằng một đoạn mỏng có viền dưới, thụt hai bên để độ dài = tỉ lệ × độ rộng dòng chữ (ước lượng theo
        /// bảng độ rộng ký tự Times New Roman). Add-in trên Windows đo chính xác vị trí chữ và vẽ shape.
        /// </summary>
        public static void InsertLineAfter(Paragraph p, ParagraphInfo info, float ratio, float availablePt, float sizePt, bool bold)
        {
            var text = TextUtil.Clean(info.Text);
            float w = TextWidth.Measure(text, sizePt, bold);
            if (w > availablePt) { w %= availablePt; if (w < availablePt * 0.15f) w = availablePt; }
            float line = w * ratio;
            int side = (int)Math.Max(0, (availablePt - line) / 2 * TwipPerPt);
            var lp = new Paragraph(new ParagraphProperties(
                new ParagraphBorders(new BottomBorder { Val = BorderValues.Single, Size = 6, Space = 0, Color = "000000" }),
                new SpacingBetweenLines { Before = "0", After = "0", Line = "40", LineRule = LineSpacingRuleValues.Exact },
                new Indentation { Left = side.ToString(), Right = side.ToString(), FirstLine = "0" },
                new Justification { Val = JustificationValues.Center },
                new ParagraphMarkRunProperties(new FontSize { Val = "4" })));
            p.InsertAfterSelf(lp);
        }
    }

    /// <summary>Độ rộng ký tự Times New Roman (đơn vị 1/1000 em, theo AFM Times-Roman / Times-Bold).</summary>
    public static class TextWidth
    {
        const string Upper = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";
        static readonly int[] UpReg = { 722, 667, 667, 722, 611, 556, 722, 722, 333, 389, 722, 611, 889, 722, 722, 556, 722, 667, 556, 611, 722, 722, 944, 722, 722, 611 };
        static readonly int[] UpBold = { 722, 667, 722, 722, 667, 611, 778, 778, 389, 500, 778, 667, 944, 722, 778, 611, 778, 722, 556, 667, 722, 722, 1000, 722, 722, 667 };
        static readonly int[] LoReg = { 444, 500, 444, 500, 444, 333, 500, 500, 278, 278, 500, 278, 778, 500, 500, 500, 500, 333, 389, 278, 500, 500, 722, 500, 500, 444 };
        static readonly int[] LoBold = { 500, 556, 444, 556, 444, 333, 500, 556, 278, 333, 556, 278, 833, 556, 500, 556, 556, 444, 389, 333, 556, 500, 722, 500, 500, 444 };

        public static float Measure(string s, float sizePt, bool bold)
        {
            float total = 0;
            foreach (var ch0 in TextUtil.StripDiacritics(s))
            {
                char ch = ch0;
                int w;
                int u = Upper.IndexOf(ch);
                if (u >= 0) w = bold ? UpBold[u] : UpReg[u];
                else if (ch >= 'a' && ch <= 'z') w = bold ? LoBold[ch - 'a'] : LoReg[ch - 'a'];
                else if (char.IsDigit(ch)) w = 500;
                else if (ch == ' ' || ch == ',' || ch == '.') w = 250;
                else if (ch == '-' || ch == ':' || ch == ';') w = 333;
                else if (ch == '/') w = 278;
                else w = 500;
                total += w;
            }
            return total / 1000f * sizePt;
        }
    }
}
