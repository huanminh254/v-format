using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using VFormat.Core;

namespace VFormat.Docx
{
    public class FormatOptions
    {
        public StandardKind? Standard;     // null = tự nhận diện
        public float BodySize = 14;
        public float LineSpacing = 1.15f;
    }

    public class FormatResult
    {
        public string Input, Output, Report;
        public FormatProfile Profile;
        public List<Issue> IssuesBefore, IssuesAfter;
        public List<string> Stats;
        public int SchemaErrors;
    }

    /// <summary>Chuẩn hóa tệp .docx không cần Microsoft Word (dùng cùng lõi VFormat.Core với add-in).</summary>
    public static class DocxFormatter
    {
        public static FormatResult Run(string input, string output, FormatOptions options = null)
        {
            options = options ?? new FormatOptions();
            if (!input.EndsWith(".docx", StringComparison.OrdinalIgnoreCase))
                throw new NotSupportedException("Chỉ hỗ trợ tệp .docx. Hãy mở tệp .doc bằng Word và lưu lại dạng .docx.");
            if (Path.GetFullPath(input) != Path.GetFullPath(output)) File.Copy(input, output, overwrite: true);
            FormatProfile p;
            ClassifiedDocument before;
            List<Issue> issuesBefore;
            var stats = new List<string>();

            using (var doc = WordprocessingDocument.Open(output, true))
            {
                var m = DocxModel.Read(doc);
                var kind = options.Standard ?? StandardDetector.Detect(m.Infos);
                p = Profiles.Get(kind).Clone();
                if (kind != StandardKind.Dang_HD05) p.SetBodySize(options.BodySize);
                p.LineSpacing = options.LineSpacing;

                before = Classifier.Classify(m.Infos, p.Kind);
                issuesBefore = FormatChecker.Check(before, p, DocxWriter.ReadPage(doc));

                DocxWriter.PageSetup(doc, p);
                var cd = before;
                var hp = Planner.PlanHeader(cd, p);
                if (hp != null)
                {
                    DocxWriter.ReplaceWithTable(m, hp); stats.Add("dàn khối đầu thành bảng 2 cột");
                    m = DocxModel.Read(doc); cd = Classifier.Classify(m.Infos, p.Kind);
                }
                var sp = Planner.PlanSign(cd, p);
                if (sp != null)
                {
                    DocxWriter.ReplaceWithTable(m, sp); stats.Add("dàn khối chữ ký & nơi nhận thành bảng 2 cột");
                    m = DocxModel.Read(doc); cd = Classifier.Classify(m.Infos, p.Kind);
                }

                int fixes = 0;
                int lastNoiNhan = cd.Kinds.LastIndexOf(ElementKind.NoiNhanItem);
                for (int i = 0; i < cd.Kinds.Count; i++)
                {
                    if (!Planner.IsFormal(cd.Kinds[i])) continue;
                    var raw = m.Infos[i].Text;
                    var clean = TextUtil.Clean(raw);
                    var fixedText = TextNormalizer.Fix(cd.Kinds[i], clean, p, i == lastNoiNhan);
                    if (fixedText.Length == 0 || fixedText == raw) continue;
                    DocxWriter.SetText(m.Paras[i], fixedText); fixes++;
                }
                var (first, last) = Planner.BodyRange(cd);
                var (headerT, signT) = Planner.FormalTables(cd);
                for (int i = Math.Max(first, 0); first >= 0 && i <= last; i++)
                    if (!headerT.Contains(m.Infos[i].TableIndex) && !signT.Contains(m.Infos[i].TableIndex))
                        fixes += DocxWriter.FixSpacing(m.Paras[i]);
                stats.Add($"{fixes} chỗ sửa chữ");

                m = DocxModel.Read(doc); cd = Classifier.Classify(m.Infos, p.Kind);
                for (int i = 0; i < cd.Kinds.Count; i++)
                    DocxWriter.Apply(m.Paras[i], Planner.PlanParagraph(cd.Kinds[i], m.Infos[i], p,
                        inList: m.Paras[i].ParagraphProperties?.NumberingProperties != null), p.FontName);
                stats.Add($"{DocxWriter.FormatTables(m, cd, p)} bảng");
                DocxWriter.PageNumbers(doc, p);

                (headerT, signT) = Planner.FormalTables(cd);
                var lines = Planner.PlanLines(cd, p);
                foreach (var (i, ratio) in lines.OrderByDescending(x => x.index))
                {
                    var info = m.Infos[i];
                    float avail = !info.InTable ? p.TextWidthMm
                        : headerT.Contains(info.TableIndex) ? (info.CellColumn <= 1 ? p.HeaderLeftColMm : p.HeaderRightColMm)
                        : signT.Contains(info.TableIndex) ? (info.CellColumn <= 1 ? p.SignLeftColMm : p.SignRightColMm) : p.TextWidthMm;
                    var st = p.StyleFor(cd.Kinds[i]);
                    DocxWriter.InsertLineAfter(m.Paras[i], info, ratio, avail * 72f / 25.4f, st.Size, st.Bold);
                }
                stats.Add($"{lines.Count} đường kẻ");
                doc.Save();
            }

            // Rà soát lại đầu ra
            ClassifiedDocument after;
            List<Issue> issuesAfter;
            int schemaErrors;
            using (var doc = WordprocessingDocument.Open(output, false))
            {
                var m = DocxModel.Read(doc);
                after = Classifier.Classify(m.Infos, p.Kind);
                issuesAfter = FormatChecker.Check(after, p, DocxWriter.ReadPage(doc));
                var schema = new DocumentFormat.OpenXml.Validation.OpenXmlValidator(DocumentFormat.OpenXml.FileFormatVersions.Office2013)
                    .Validate(doc).ToList();
                schemaErrors = schema.Count;
                stats.Add(schema.Count == 0 ? "OOXML hợp lệ (schema Office 2013)"
                    : $"{schema.Count} LỖI SCHEMA: " + string.Join("; ", schema.Take(3).Select(e => e.Description + " @ " + e.Path?.XPath)));
            }
            return new FormatResult
            {
                Input = input, Output = output, Profile = p, IssuesBefore = issuesBefore, IssuesAfter = issuesAfter, Stats = stats,
                SchemaErrors = schemaErrors,
                Report = Report(input, output, p, before, issuesBefore, after, issuesAfter, stats),
            };
        }

        /// <summary>Tạo văn bản mới từ mẫu trong thư viện rồi chuẩn hóa.</summary>
        public static FormatResult CreateFromTemplate(DocTemplate t, string output, FormatOptions options = null, string diaDanh = null)
        {
            var tmp = Path.Combine(Path.GetTempPath(), "vformat-" + Guid.NewGuid().ToString("N") + ".docx");
            using (var doc = WordprocessingDocument.Create(tmp, DocumentFormat.OpenXml.WordprocessingDocumentType.Document))
            {
                var main = doc.AddMainDocumentPart();
                var body = new Body();
                foreach (var line in t.Lines)
                {
                    var text = diaDanh != null ? line.Replace("Hà Nội, ngày", diaDanh + ", ngày") : line;
                    body.AppendChild(new Paragraph(new Run(new Text(text) { Space = DocumentFormat.OpenXml.SpaceProcessingModeValues.Preserve })));
                }
                body.AppendChild(new SectionProperties());
                main.Document = new Document(body);
            }
            try
            {
                var o = options ?? new FormatOptions();
                return Run(tmp, output, new FormatOptions { Standard = t.Standard, BodySize = o.BodySize, LineSpacing = o.LineSpacing });
            }
            finally { try { File.Delete(tmp); } catch { } }
        }

        static string Report(string input, string output, FormatProfile p, ClassifiedDocument before, List<Issue> ib,
            ClassifiedDocument after, List<Issue> ia, List<string> stats)
        {
            var sb = new StringBuilder();
            sb.AppendLine($"# {Path.GetFileName(input)} → {Path.GetFileName(output)}");
            sb.AppendLine();
            sb.AppendLine($"**Quy chuẩn:** {p.FullName} — {p.Description}  ");
            sb.AppendLine($"**Đã làm:** {string.Join(", ", stats)}  ");
            sb.AppendLine($"**Lỗi thể thức:** trước {ib.Count} → sau {ia.Count}");
            sb.AppendLine();
            sb.AppendLine("## Đầu vào — nhận diện thành phần");
            sb.AppendLine();
            Table(sb, before);
            sb.AppendLine("## Lỗi phát hiện ở đầu vào");
            sb.AppendLine();
            Issues(sb, ib, before);
            sb.AppendLine("## Đầu ra — nhận diện thành phần");
            sb.AppendLine();
            Table(sb, after);
            sb.AppendLine("## Lỗi còn lại ở đầu ra");
            sb.AppendLine();
            Issues(sb, ia, after);
            return sb.ToString();
        }

        static void Table(StringBuilder sb, ClassifiedDocument cd)
        {
            sb.AppendLine("| # | Bảng | Thành phần | Nội dung | Phông / cỡ / B I |");
            sb.AppendLine("|---|---|---|---|---|");
            for (int i = 0; i < cd.Kinds.Count; i++)
            {
                if (cd.Kinds[i] == ElementKind.Empty) continue;
                var x = cd.Paragraphs[i];
                var text = TextUtil.Clean(x.Text).Replace("|", "\\|");
                if (text.Length > 70) text = text.Substring(0, 67) + "...";
                var raw = x.Text.Replace("\t", " ");
                if (raw.Contains("  ")) text = text + " ⟨khoảng trắng thừa⟩";
                var fmt = x.FontName == null ? "" : $"{x.FontName} {x.FontSize:0.#} {(x.Bold == 1 ? "B" : x.Bold == 0 ? "-" : "?")}{(x.Italic == 1 ? "I" : x.Italic == 0 ? "-" : "?")}";
                sb.AppendLine($"| {i} | {(x.InTable ? $"B{x.TableIndex}·c{x.CellColumn}" : "")} | {FormatChecker.Describe(cd.Kinds[i])} | {text} | {fmt} |");
            }
            sb.AppendLine();
        }

        static void Issues(StringBuilder sb, List<Issue> issues, ClassifiedDocument cd)
        {
            if (issues.Count == 0) { sb.AppendLine("_Không có._"); sb.AppendLine(); return; }
            foreach (var i in issues)
                sb.AppendLine($"- `{i.Code}` {(i.ParagraphIndex >= 0 ? $"(đoạn {i.ParagraphIndex}) " : "")}{i.Message}");
            sb.AppendLine();
        }
    }
}
