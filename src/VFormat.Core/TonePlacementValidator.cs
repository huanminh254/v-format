using System.Collections.Generic;
using System.Text;
using System.Text.RegularExpressions;

namespace VFormat.Core
{
    /// <summary>
    /// Kiểm tra và sửa vị trí đặt dấu thanh tiếng Việt theo kiểu mới
    /// (QĐ 1989/QĐ-BGDĐT Điều 9: "hòa", "thúy" thay vì "hoà", "thuý"), hoặc chuyển sang kiểu cũ.
    /// Dựa trên cấu trúc âm tiết: phụ âm đầu (kể cả qu, gi) + cụm nguyên âm + âm cuối.
    /// </summary>
    public static class TonePlacementValidator
    {
        // Nguyên âm có thanh: 0=thanh ngang, 1=huyền, 2=sắc, 3=hỏi, 4=ngã, 5=nặng
        static readonly Dictionary<char, (char baseChar, int tone)> ToneMap;
        static readonly Dictionary<string, char> ReverseToneMap;

        // Nguyên âm có mũ/móc: nếu âm tiết có, dấu thanh luôn đặt ở đó (ươ → ơ).
        const string MarkedVowels = "ăâêôơư";
        static readonly HashSet<string> Codas = new HashSet<string> { "c", "ch", "m", "n", "ng", "nh", "p", "t" };

        static TonePlacementValidator()
        {
            ToneMap = new Dictionary<char, (char, int)>();
            ReverseToneMap = new Dictionary<string, char>();
            Register('a', "aàáảãạ");
            Register('ă', "ăằắẳẵặ");
            Register('â', "âầấẩẫậ");
            Register('e', "eèéẻẽẹ");
            Register('ê', "êềếểễệ");
            Register('i', "iìíỉĩị");
            Register('o', "oòóỏõọ");
            Register('ô', "ôồốổỗộ");
            Register('ơ', "ơờớởỡợ");
            Register('u', "uùúủũụ");
            Register('ư', "ưừứửữự");
            Register('y', "yỳýỷỹỵ");
        }

        static void Register(char baseChar, string group)
        {
            for (int i = 0; i < group.Length; i++)
            {
                var c = group[i];
                ToneMap[c] = (baseChar, i);
                ReverseToneMap[baseChar + "_" + i] = c;
            }
        }

        /// <summary>Tách dấu thanh khỏi âm tiết: trả về tone (0..5), âm tiết không dấu, vị trí nguyên âm có dấu.</summary>
        public static int ExtractTone(string word, out string baseWord, out int tonePos)
        {
            tonePos = -1; int tone = 0;
            var sb = new StringBuilder();
            for (int i = 0; i < word.Length; i++)
            {
                var ch = word[i];
                var low = char.ToLowerInvariant(ch);
                if (ToneMap.TryGetValue(low, out var v))
                {
                    if (v.tone > 0) { tone = v.tone; tonePos = i; }
                    sb.Append(char.IsUpper(ch) ? char.ToUpperInvariant(v.baseChar) : v.baseChar);
                }
                else sb.Append(ch);
            }
            baseWord = sb.ToString();
            return tone;
        }

        /// <summary>Đặt dấu thanh vào vị trí vowelIndex của âm tiết baseWord.</summary>
        public static string ApplyTone(string baseWord, int tone, int vowelIndex)
        {
            if (tone <= 0 || vowelIndex < 0 || vowelIndex >= baseWord.Length) return baseWord;
            var ch = baseWord[vowelIndex];
            var low = char.ToLowerInvariant(ch);
            if (!ReverseToneMap.TryGetValue(low + "_" + tone, out var toned)) return baseWord;
            var sb = new StringBuilder(baseWord);
            sb[vowelIndex] = char.IsUpper(ch) ? char.ToUpperInvariant(toned) : toned;
            return sb.ToString();
        }

        /// <summary>
        /// Vị trí đặt dấu thanh đúng của âm tiết baseWord (đã bỏ dấu thanh), -1 nếu không phải một âm tiết tiếng Việt
        /// (từ dính liền, từ nước ngoài...). newStyle: oa/oe/uy không có âm cuối đặt dấu ở nguyên âm đầu (hòa, thủy);
        /// kiểu cũ đặt ở nguyên âm sau (hoà, thuỷ).
        /// </summary>
        static int TonePosition(string baseWord, bool newStyle)
        {
            var w = baseWord.ToLowerInvariant();
            bool IsVowel(int i) => i < w.Length && ToneMap.ContainsKey(w[i]);

            int start = 0;
            while (start < w.Length && !IsVowel(start)) start++;
            if (start == w.Length) return -1;
            int end = start;
            while (end < w.Length && IsVowel(end)) end++;
            for (int i = end; i < w.Length; i++) if (IsVowel(i)) return -1;   // hai cụm nguyên âm: từ dính liền
            for (int i = 0; i < start; i++) if (!char.IsLetter(w[i])) return -1;
            var coda = w.Substring(end);
            if (coda.Length > 0 && !Codas.Contains(coda)) return -1;

            // "qu", "gi" là phụ âm đầu: u/i không phải nguyên âm (quả, quyết, giữa) — trừ "gì", "gìn".
            if (start > 0 && end - start >= 2 &&
                ((w[start - 1] == 'q' && w[start] == 'u') || (start == 1 && w[0] == 'g' && w[start] == 'i')))
                start++;

            int n = end - start;
            if (n == 1) return start;
            if (n > 3) return -1;

            int marked = -1;
            for (int i = start; i < end; i++)
                if (MarkedVowels.IndexOf(w[i]) >= 0) marked = i;     // ươ: lấy ơ (đứng sau)
            if (marked >= 0) return marked;

            if (coda.Length > 0) return end - 1;                     // hoàng, toán, huýt
            if (n == 3) return start + 1;                            // loại, khuỷu, ngoẻo
            var pair = w.Substring(start, 2);
            if (pair == "oa" || pair == "oe" || pair == "uy") return newStyle ? start : start + 1;
            return start;                                            // tài, cháo, mùa, kìa, túi
        }

        /// <summary>Đặt lại dấu thanh của một âm tiết; trả về nguyên từ nếu không có dấu hoặc không xác định được.</summary>
        public static string PlaceTone(string word, bool newStyle = true)
        {
            if (string.IsNullOrEmpty(word)) return word;
            int tone = ExtractTone(word, out var baseWord, out var tonePos);
            if (tone == 0) return word;
            int marks = 0;
            foreach (var ch in word)
                if (ToneMap.TryGetValue(char.ToLowerInvariant(ch), out var v) && v.tone > 0) marks++;
            if (marks > 1) return word;
            int pos = TonePosition(baseWord, newStyle);
            if (pos < 0 || pos == tonePos) return word;
            return ApplyTone(baseWord, tone, pos);
        }

        /// <summary>
        /// Kiểm tra âm tiết đặt dấu thanh đúng chỗ (kiểu mới: hòa, thủy). Trả về true nếu đúng;
        /// nếu sai, correctedWord chứa phiên bản đúng.
        /// </summary>
        public static bool IsTonePlacedProperly(string word, out string correctedWord)
        {
            correctedWord = PlaceTone(word, newStyle: true);
            return correctedWord == word;
        }

        /// <summary>
        /// Trả về biến thể vị trí dấu kiểu còn lại của từ (ví dụ "hòa" ↔ "hoà"), null nếu từ không có hai cách đặt dấu.
        /// </summary>
        public static string GetAlternativeTonePlacement(string word)
        {
            if (string.IsNullOrEmpty(word)) return null;
            word = word.Normalize(NormalizationForm.FormC);
            var newer = PlaceTone(word, true);
            var older = PlaceTone(word, false);
            if (newer == older) return null;
            if (word == newer) return older;
            if (word == older) return newer;
            return null;
        }

        /// <summary>Trả về chuỗi với mọi âm tiết đặt dấu lại theo kiểu mới (newStyle) hoặc kiểu cũ.</summary>
        public static string NormalizeAll(string text, bool newStyle = true)
        {
            if (string.IsNullOrEmpty(text)) return text ?? "";
            text = text.Normalize(NormalizationForm.FormC);
            return Regex.Replace(text, @"[\p{L}\p{M}]+", m => PlaceTone(m.Value, newStyle));
        }
    }
}
