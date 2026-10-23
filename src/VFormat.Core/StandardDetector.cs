using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

namespace VFormat.Core
{
    /// <summary>Tự nhận diện hệ quy chuẩn dựa vào khối đầu và tên loại văn bản.</summary>
    public static class StandardDetector
    {
        static readonly string[] MilitaryKeys =
        {
            "BO QUOC PHONG", "QUAN KHU", "QUAN DOAN", "SU DOAN", "LU DOAN", "TRUNG DOAN", "TIEU DOAN",
            "BO CHI HUY QUAN SU", "BAN CHI HUY QUAN SU", "BO TU LENH", "QUAN CHUNG", "BINH CHUNG",
            "TONG CUC CHINH TRI", "BO TONG THAM MUU", "HOC VIEN LUC QUAN", "BO DOI BIEN PHONG",
            "CANH SAT BIEN", "TONG CUC HAU CAN", "TONG CUC KY THUAT", "QUAN DOI NHAN DAN",
        };

        public static StandardKind Detect(IList<ParagraphInfo> paragraphs)
        {
            var head = paragraphs.Take(25).Select(p => TextUtil.Key(p.Text)).Where(t => t.Length > 0).ToList();
            var all = paragraphs.Take(60).Select(p => TextUtil.Key(p.Text)).ToList();

            if (all.Any(t => t.Contains("KE HOACH BAI DAY") || t.StartsWith("GIAO AN")))
                return StandardKind.GiaoDuc_CV5512;
            if (head.Any(t => t == "DANG CONG SAN VIET NAM"))
                return StandardKind.Dang_HD05;
            if (head.Any(t => Regex.IsMatch(t, @"^DOAN (TNCS|THANH NIEN CONG SAN) HO CHI MINH$")))
                return StandardKind.Doan_HD29;
            if (all.Any(t => t.StartsWith("VAN BAN HOP NHAT") || t.Contains("PHAP DIEN")))
                return StandardKind.PhapDien_ND250;
            if (head.Any(t => MilitaryKeys.Any(k => t.Contains(k))))
                return StandardKind.QuanDoi_TT68;
            return StandardKind.HanhChinh_ND30;
        }
    }
}
