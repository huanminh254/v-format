using System.Text.RegularExpressions;

namespace VFormat.Core
{
    /// <summary>Sửa nội dung chữ của các thành phần thể thức. Trả về chuỗi mới (giống chuỗi cũ nếu không cần sửa).</summary>
    public static class TextNormalizer
    {
        public const string QuocHieuText = "CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM";
        public const string TieuNguText = "Độc lập - Tự do - Hạnh phúc";

        /// <param name="isLast">đoạn cuối cùng cùng loại (nơi nhận cuối, căn cứ cuối) — kết thúc bằng dấu chấm</param>
        public static string Fix(ElementKind kind, string text, FormatProfile profile, bool isLast = false)
        {
            var s = TextUtil.Clean(text);
            if (s.Length == 0) return s;
            switch (kind)
            {
                case ElementKind.QuocHieu: return QuocHieuText;
                case ElementKind.TieuNgu: return TieuNguText;
                case ElementKind.TieuDeDang: return "ĐẢNG CỘNG SẢN VIỆT NAM";
                case ElementKind.TieuDeDoan: return "ĐOÀN TNCS HỒ CHÍ MINH";
                case ElementKind.SoKyHieu: return FixSoKyHieu(s, profile);
                case ElementKind.DiaDanhNgayThang: return FixNgayThang(s);
                case ElementKind.TrichYeuCongVan: return Regex.Replace(s, @"^V\s*/\s*v\s*:?\s*", "V/v ", RegexOptions.IgnoreCase);
                case ElementKind.KinhGui: return Regex.Replace(s, @"^Kính\s+gửi\s*:?\s*", "Kính gửi: ", RegexOptions.IgnoreCase).TrimEnd();
                case ElementKind.NoiNhanTitle: return "Nơi nhận:";
                case ElementKind.NoiNhanItem: return FixNoiNhanItem(s, isLast, profile.NoiNhanSeparator);
                case ElementKind.CanCu: return profile.CanCuDash ? FixCanCuDash(s, isLast) : FixSpacing(s);
                case ElementKind.Dieu: return Regex.Replace(s, @"^Điều\s+(\d+[a-z]?)\s*[\.:]\s*", "Điều $1. ");
                case ElementKind.Phan: case ElementKind.Chuong:
                    // số thứ tự La Mã viết hoa: "Chương i" → "Chương I"
                    return Regex.Replace(FixSpacing(s), @"^(Phần|Chương)\s+([ivxlcIVXLC]+)\b", m => m.Groups[1].Value + " " + m.Groups[2].Value.ToUpperInvariant(), RegexOptions.IgnoreCase);
                case ElementKind.TrichYeu when profile.Kind == StandardKind.GiaoDuc_CV5512:
                    // KHBD (CV 5512 Phụ lục IV): "TÊN BÀI DẠY: ..."
                    return Regex.Replace(FixSpacing(s), @"^tên\s+bài\s+dạy\s*:\s*", "TÊN BÀI DẠY: ", RegexOptions.IgnoreCase);
                case ElementKind.ChucVuKy: return FixChucVu(s, profile);
                case ElementKind.HoTenKy: return s;
                default: return FixSpacing(s);
            }
        }

        /// <summary>"Số :5 /QĐ- UBND" → "Số: 05/QĐ-UBND"; Đảng: "Số 05-NQ/TU". Số nhỏ hơn 10 thêm số 0 phía trước.</summary>
        public static string FixSoKyHieu(string s, FormatProfile p)
        {
            var m = Regex.Match(s, @"^Số\s*:?\s*(.*)$", RegexOptions.IgnoreCase);
            if (!m.Success) return s;
            var prefix = p.SoKyHieuHasColon ? "Số:" : "Số";
            // chưa có số: giữ khoảng trống để điền số ("Số:      /QĐ-UBND")
            var blank = Regex.Match(m.Groups[1].Value, @"^[\s\.…_]*([/\-].*)$");
            if (blank.Success) return prefix + "      " + Compact(blank.Groups[1].Value);
            var rest = Regex.Replace(Compact(m.Groups[1].Value), @"^(\d)(?=[/\-])", "0$1");
            return prefix + " " + rest;
        }

        /// <summary>Đảng (HD 05 mục 6.2.3): "- Căn cứ ...;", căn cứ cuối cùng kết thúc bằng dấu chấm.</summary>
        public static string FixCanCuDash(string s, bool isLast)
        {
            s = FixSpacing(Regex.Replace(s, @"^[\-–+•]\s*", ""));
            s = Regex.Replace(s, @"[\s;,\.]+$", "");
            return "- " + s + (isLast ? "." : ";");
        }

        static string Compact(string x) => Regex.Replace(x, @"\s*([/\-])\s*", "$1").Trim();

        /// <summary>
        /// Ngày nhỏ hơn 10 và tháng 1, 2 phải thêm số 0 phía trước (NĐ 30, Phụ lục I).
        /// "Hà Nội, ngày 5 tháng 1 năm 2026" → "Hà Nội, ngày 05 tháng 01 năm 2026".
        /// </summary>
        public static string FixNgayThang(string s)
        {
            s = Regex.Replace(s, @"\s*,\s*ngày", ", ngày", RegexOptions.IgnoreCase);
            s = Regex.Replace(s, @"ngày\s+(\d{1,2})\s+tháng\s+(\d{1,2})\s+năm\s+(\d{4})", m =>
            {
                int d = int.Parse(m.Groups[1].Value), mo = int.Parse(m.Groups[2].Value);
                var ds = d < 10 ? d.ToString("00") : d.ToString();
                var ms = mo <= 2 ? mo.ToString("00") : mo.ToString();
                return $"ngày {ds} tháng {ms} năm {m.Groups[3].Value}";
            }, RegexOptions.IgnoreCase);
            return Regex.Replace(s, @"[\s\.]+$", "");
        }

        public static string FixNoiNhanItem(string s, bool isLast, string separator = ";")
        {
            s = FixSpacing(s);   // "- Lưu : VT" → "- Lưu: VT"
            s = Regex.Replace(s, @"^[\-\+•–]\s*", "- ");
            if (!s.StartsWith("- ")) s = "- " + s;
            s = Regex.Replace(s, @"[\s;,]+$", "");
            // bỏ dấu chấm kết thúc cũ nhưng giữ "..." (chỗ để trống): "Lưu: VT." → "Lưu: VT", "VT, ...." → "VT, ..."
            int dots = s.Length - s.TrimEnd('.').Length;
            if (dots % 3 == 1) s = s.Substring(0, s.Length - 1).TrimEnd();
            return s + (isLast ? "." : separator);
        }

        static string FixChucVu(string s, FormatProfile p)
        {
            // ký hiệu thẩm quyền ký theo quy chuẩn: NĐ 30 "TM./KT./TL./TUQ.", Đảng "T/M, K/T, T/L, T/UQ"
            s = Regex.Replace(s, @"^(T\.?M\.?|T/M)\s+", p.ThayMatText + " ", RegexOptions.IgnoreCase);
            s = Regex.Replace(s, @"^(K\.?T\.?|K/T)\s+", p.KyThayText + " ", RegexOptions.IgnoreCase);
            s = Regex.Replace(s, @"^(T\.?L\.?|T/L)\s+", p.ThuaLenhText + " ", RegexOptions.IgnoreCase);
            s = Regex.Replace(s, @"^(TUQ\.?|T/UQ)\s+", p.ThuaUyQuyenText + " ", RegexOptions.IgnoreCase);
            return TextUtil.Upper(s);
        }

        /// <summary>Bỏ khoảng trắng thừa, khoảng trắng trước dấu câu, thêm khoảng trắng sau dấu câu.</summary>
        public static string FixSpacing(string s)
        {
            s = Regex.Replace(s, @" {2,}", " ");
            s = Regex.Replace(s, @" +([,;:\?!\)\]]|\.(?!\.))", "$1");   // giữ " ..." (chỗ để trống)
            s = Regex.Replace(s, @"([\(\[]) +", "$1");
            s = Regex.Replace(s, @"(\d) +/ *(?=[A-ZĐ\d])", "$1/");   // "số 15 /KH-VP" → "số 15/KH-VP"
            s = Regex.Replace(s, @"([,;])(?=[^\s\d""”’\)])", "$1 ");
            return s.Trim();
        }
    }
}
