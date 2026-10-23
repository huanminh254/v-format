using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml.Packaging;
using VFormat.Core;

namespace VFormat.Cli
{
    /// <summary>
    /// vformat — chuẩn hóa .docx không cần Word, dùng cùng lõi VFormat.Core với add-in.
    ///   vformat input.docx output.docx [--standard nd30|tt68|hd05|hd29|nd250|cv5512] [--report report.md]
    /// </summary>
    public static class Program
    {
        public static int Main(string[] args)
        {
            Console.OutputEncoding = Encoding.UTF8;
            var pos = args.Where(a => !a.StartsWith("--")).ToList();
            string Opt(string name) { int i = Array.IndexOf(args, name); return i >= 0 && i + 1 < args.Length ? args[i + 1] : null; }
            if (pos.Count < 2)
            {
                Console.Error.WriteLine("Cách dùng: vformat <vào.docx> <ra.docx> [--standard nd30|tt68|hd05|hd29|nd250|cv5512] [--report báo-cáo.md]");
                return 2;
            }
            pos.Remove(Opt("--standard")); pos.Remove(Opt("--report"));
            StandardKind? force = null;
            switch (Opt("--standard")?.ToLowerInvariant())
            {
                case "nd30": force = StandardKind.HanhChinh_ND30; break;
                case "tt68": force = StandardKind.QuanDoi_TT68; break;
                case "hd05": force = StandardKind.Dang_HD05; break;
                case "hd29": force = StandardKind.Doan_HD29; break;
                case "nd250": force = StandardKind.PhapDien_ND250; break;
                case "cv5512": force = StandardKind.GiaoDuc_CV5512; break;
            }
            var report = Run(pos[0], pos[1], force);
            var reportPath = Opt("--report");
            if (reportPath != null) File.WriteAllText(reportPath, report, new UTF8Encoding(false));
            else Console.WriteLine(report);
            return 0;
        }

        public static string Run(string input, string output, StandardKind? force)
        {
            File.Copy(input, output, overwrite: true);
            FormatProfile p;
            ClassifiedDocument before;
            List<Issue> issuesBefore;
            var stats = new List<string>();

            using (var doc = WordprocessingDocument.Open(output, true))
            {
                var m = DocxModel.Read(doc);
                var kind = force ?? StandardDetector.Detect(m.Infos);
                p = Profiles.Get(kind).Clone();
                if (kind != StandardKind.Dang_HD05) p.SetBodySize(14);
                p.LineSpacing = 1.15f;

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
            using (var doc = WordprocessingDocument.Open(output, false))
            {
                var m = DocxModel.Read(doc);
                after = Classifier.Classify(m.Infos, p.Kind);
                issuesAfter = FormatChecker.Check(after, p, DocxWriter.ReadPage(doc));
                var schema = new DocumentFormat.OpenXml.Validation.OpenXmlValidator(DocumentFormat.OpenXml.FileFormatVersions.Office2013)
                    .Validate(doc).ToList();
                stats.Add(schema.Count == 0 ? "OOXML hợp lệ (schema Office 2013)"
                    : $"{schema.Count} LỖI SCHEMA: " + string.Join("; ", schema.Take(3).Select(e => e.Description + " @ " + e.Path?.XPath)));
            }
            return Report(input, output, p, before, issuesBefore, after, issuesAfter, stats);
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
