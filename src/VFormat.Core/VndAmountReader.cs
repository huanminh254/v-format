using System.Collections.Generic;
using System.Text;
using System.Text.RegularExpressions;

namespace VFormat.Core
{
    /// <summary>Tuỳ chọn đọc số tiền VND thành chữ.</summary>
    public class VndReadingOptions
    {
        public bool CapitalizeFirst = true;   // viết hoa chữ đầu
        public bool ZeroTensAsLe = false;     // "linh" (mặc định) hay "lẻ" cho hàng chục bằng 0
        public bool ThousandAsNgan = false;   // "nghìn" (mặc định) hay "ngàn"
        public bool FourAsTu = false;         // "bốn" (mặc định) hay "tư" cho hàng đơn vị bằng 4 sau mười
        public static readonly VndReadingOptions Default = new VndReadingOptions();
    }

    /// <summary>
    /// Chuyển số tiền VND thành chữ tiếng Việt (ví dụ "2.345.000 đồng" → "Hai triệu ba trăm bốn mươi lăm nghìn đồng").
    /// Chấp nhận số nguyên tối đa 18 chữ số; không có phần lẻ thập phân.
    /// </summary>
    public static class VndAmountReader
    {
        public const int MaxDigits = 18;

        static readonly string[] Digits = { "không", "một", "hai", "ba", "bốn", "năm", "sáu", "bảy", "tám", "chín" };

        static readonly Regex CurrencySuffix = new Regex(@"\s*(vnđ|vnd|đồng|đ|₫)\s*\.?\s*$", RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);
        static readonly Regex DotGrouped = new Regex(@"^\d{1,3}(?:\.\d{3})+(?:,(\d+))?$");
        static readonly Regex CommaGrouped = new Regex(@"^\d{1,3}(?:,\d{3})+(?:\.(\d+))?$");
        static readonly Regex SpaceGrouped = new Regex(@"^\d{1,3}(?: \d{3})+$");
        static readonly Regex Plain = new Regex(@"^\d+(?:[.,](\d+))?$");

        public static bool TryConvert(string text, out string words, out string error)
            => TryConvert(text, null, out words, out error);

        public static bool TryConvert(string text, VndReadingOptions options, out string words, out string error)
        {
            words = null;
            if (!TryParse(text, out var digits, out var negative, out error)) return false;
            words = ToWords(digits, negative, options);
            return true;
        }

        public static bool TryParse(string text, out string digits, out bool negative, out string error)
        {
            digits = null; negative = false; error = null;
            var input = (text ?? "").Replace(' ', ' ').Replace(' ', ' ').Trim();
            input = CurrencySuffix.Replace(input, "").Trim();
            if (input.StartsWith("-") || input.StartsWith("−"))
            {
                negative = true;
                input = input.Substring(1).Trim();
            }
            var m = DotGrouped.Match(input);
            if (!m.Success) m = CommaGrouped.Match(input);
            if (!m.Success) m = SpaceGrouped.Match(input);
            if (!m.Success) m = Plain.Match(input);
            if (!m.Success)
            {
                error = "Vùng chọn không phải là một số tiền (ví dụ hợp lệ: 2345, 2.345.000, 2.345.000 đồng).";
                return false;
            }
            var intPart = input;
            if (m.Groups.Count > 1 && m.Groups[1].Success)
            {
                var frac = m.Groups[1].Value;
                if (frac.Trim('0').Length > 0)
                {
                    error = "Số tiền VND không có phần lẻ thập phân (ví dụ 2.345,50). Vui lòng làm tròn trước khi chuyển.";
                    return false;
                }
                intPart = input.Substring(0, input.Length - frac.Length - 1);
            }
            var sb = new StringBuilder(intPart.Length);
            foreach (var c in intPart) if (c >= '0' && c <= '9') sb.Append(c);
            var d = sb.ToString().TrimStart('0');
            if (d.Length == 0) d = "0";
            if (d.Length > MaxDigits)
            {
                error = "Số tiền quá lớn (tối đa 18 chữ số).";
                return false;
            }
            digits = d;
            return true;
        }

        public static string ToWords(string digits, bool negative) => ToWords(digits, negative, null);

        public static string ToWords(string digits, bool negative, VndReadingOptions options)
        {
            var o = options ?? VndReadingOptions.Default;
            var d = (digits ?? "0").TrimStart('0');
            string words = d.Length == 0 ? "không" : Read(d, false, o);
            if (negative && d.Length != 0) words = "âm " + words;
            if (o.CapitalizeFirst && words.Length > 0)
                words = char.ToUpper(words[0]) + words.Substring(1);
            return words + " đồng";
        }

        static string Read(string d, bool full, VndReadingOptions o)
        {
            d = d.TrimStart('0'); if (d.Length == 0) d = "0";
            if (d.Length > 9)
            {
                var high = d.Substring(0, d.Length - 9);
                var low = d.Substring(d.Length - 9);
                var s = Read(high, full, o) + " tỷ";
                if (low.Trim('0').Length > 0) s += " " + ReadUpToBillion(low, true, o);
                return s;
            }
            return ReadUpToBillion(d, full, o);
        }

        static string ReadUpToBillion(string d, bool full, VndReadingOptions o)
        {
            var scales = new[] { "", o.ThousandAsNgan ? " ngàn" : " nghìn", " triệu" };
            int pad = (d.Length + 2) / 3 * 3;
            d = d.PadLeft(pad, '0');
            int groups = d.Length / 3;
            var parts = new List<string>();
            bool nextFull = full;
            for (int i = 0; i < groups; i++)
            {
                var g = d.Substring(i * 3, 3);
                if (g == "000") continue;
                parts.Add(ReadGroup(g[0] - '0', g[1] - '0', g[2] - '0', nextFull, o) + scales[groups - 1 - i]);
                nextFull = true;
            }
            return string.Join(" ", parts);
        }

        static string ReadGroup(int h, int t, int u, bool full, VndReadingOptions o)
        {
            var parts = new List<string>();
            if (h > 0 || full) parts.Add(Digits[h] + " trăm");
            if (t == 0)
            {
                if (u > 0)
                {
                    if (h > 0 || full) parts.Add(o.ZeroTensAsLe ? "lẻ" : "linh");
                    parts.Add(Digits[u]);
                }
            }
            else if (t == 1)
            {
                parts.Add("mười");
                if (u > 0) parts.Add(u == 5 ? "lăm" : Digits[u]);
            }
            else
            {
                parts.Add(Digits[t] + " mươi");
                if (u == 1) parts.Add("mốt");
                else if (u == 4 && o.FourAsTu) parts.Add("tư");
                else if (u == 5) parts.Add("lăm");
                else if (u > 0) parts.Add(Digits[u]);
            }
            return string.Join(" ", parts);
        }
    }
}
