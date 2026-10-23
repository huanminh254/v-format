using System.Globalization;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

namespace VFormat.Core
{
    public static class TextUtil
    {
        public static readonly CultureInfo Vi = new CultureInfo("vi-VN");

        /// <summary>Chuẩn hóa chuỗi lấy từ Word: bỏ ký tự điều khiển, NFC, gộp khoảng trắng.</summary>
        public static string Clean(string s)
        {
            if (string.IsNullOrEmpty(s)) return "";
            var sb = new StringBuilder(s.Length);
            foreach (var ch in s)
            {
                if (ch == '\r' || ch == '\a' || ch == '\f') continue;
                if (ch == '\v' || ch == '\t' || ch == ' ') { sb.Append(' '); continue; }
                if (ch < ' ') continue;
                sb.Append(ch);
            }
            return Regex.Replace(sb.ToString().Normalize(NormalizationForm.FormC), @"\s+", " ").Trim();
        }

        public static string Upper(string s) => (s ?? "").ToUpper(Vi);

        /// <summary>Có chữ cái và toàn bộ chữ cái đều in hoa.</summary>
        public static bool IsAllUpper(string s)
        {
            if (string.IsNullOrEmpty(s)) return false;
            bool hasLetter = false;
            foreach (var ch in s)
            {
                if (!char.IsLetter(ch)) continue;
                hasLetter = true;
                if (char.IsLower(ch)) return false;
            }
            return hasLetter;
        }

        /// <summary>Bỏ dấu tiếng Việt để so khớp mềm.</summary>
        public static string StripDiacritics(string s)
        {
            if (string.IsNullOrEmpty(s)) return "";
            var d = s.Normalize(NormalizationForm.FormD);
            var sb = new StringBuilder(d.Length);
            foreach (var ch in d)
                if (CharUnicodeInfo.GetUnicodeCategory(ch) != UnicodeCategory.NonSpacingMark) sb.Append(ch);
            return sb.ToString().Normalize(NormalizationForm.FormC).Replace('đ', 'd').Replace('Đ', 'D');
        }

        /// <summary>Khóa so khớp: bỏ dấu, in hoa, gộp khoảng trắng.</summary>
        public static string Key(string s) => Regex.Replace(StripDiacritics(Clean(s)).ToUpperInvariant(), @"\s+", " ");

        public static int WordCount(string s) =>
            string.IsNullOrWhiteSpace(s) ? 0 : s.Split(new[] { ' ' }, System.StringSplitOptions.RemoveEmptyEntries).Length;

        /// <summary>"nguyễn văn an" → "Nguyễn Văn An".</summary>
        public static string TitleCase(string s) =>
            string.Join(" ", Clean(s).Split(' ').Select(w => w.Length == 0 ? w :
                char.ToUpper(w[0], Vi) + w.Substring(1).ToLower(Vi)));
    }
}
