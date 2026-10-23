using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using VFormat.Docx;
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
            var report = DocxFormatter.Run(pos[0], pos[1], new FormatOptions { Standard = force }).Report;
            var reportPath = Opt("--report");
            if (reportPath != null) File.WriteAllText(reportPath, report, new UTF8Encoding(false));
            else Console.WriteLine(report);
            return 0;
        }

    }
}
