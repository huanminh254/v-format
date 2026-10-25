using System.Collections.Generic;
using System.Text;
using System.Text.RegularExpressions;

namespace VFormat.Core
{
    /// <summary>
    /// Kiểm tra và sửa vị trí đặt dấu thanh tiếng Việt theo kiểu mới
    /// (QĐ 1989/QĐ-BGDĐT Điều 9: "hòa", "thúy" thay vì "hoà", "thuý").
    /// Cũng chuẩn hoá hai biến thể hoà/hòa trong NFC Unicode.
    /// </summary>
    public static class TonePlacementValidator
    {
        // Nguyên âm có thanh: 0=thanh ngang, 1=huyền, 2=sắc, 3=hỏi, 4=ngã, 5=nặng
        static readonly Dictionary<char, (char baseChar, int tone)> ToneMap;
        static readonly Dictionary<string, char> ReverseToneMap;

        // Biến thể vị trí dấu kiểu cũ ↔ kiểu mới (sau NFC). Kiểu mới ở bên trái.
        static readonly Dictionary<string, string> ToneVariantMap = new Dictionary<string, string>
        {
            {"óa","oá"}, {"oá","óa"}, {"òa","oà"}, {"oà","òa"},
            {"ỏa","oả"}, {"oả","ỏa"}, {"õa","oã"}, {"oã","õa"},
            {"ọa","oạ"}, {"oạ","ọa"},
            {"óe","oé"}, {"oé","óe"}, {"òe","oè"}, {"oè","òe"},
            {"ỏe","oẻ"}, {"oẻ","ỏe"}, {"õe","oẽ"}, {"oẽ","õe"},
            {"ọe","oẹ"}, {"oẹ","ọe"},
            {"úy","uý"}, {"uý","úy"}, {"ùy","uỳ"}, {"uỳ","ùy"},
            {"ủy","uỷ"}, {"uỷ","ủy"}, {"ũy","uỹ"}, {"uỹ","ũy"},
            {"ụy","uỵ"}, {"uỵ","ụy"},
        };

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
        /// Kiểm tra âm tiết đặt dấu thanh đúng chỗ (kiểu mới QĐ 1989). Trả về true nếu đúng;
        /// nếu sai, correctedWord chứa phiên bản đúng.
        /// </summary>
        public static bool IsTonePlacedProperly(string word, out string correctedWord)
        {
            correctedWord = word;
            if (string.IsNullOrEmpty(word)) return true;
            int tone = ExtractTone(word, out var baseWord, out var tonePos);
            if (tone == 0 || tonePos == -1) return true;

            var vowels = new List<int>();
            for (int i = 0; i < baseWord.Length; i++)
                if (ToneMap.ContainsKey(char.ToLowerInvariant(baseWord[i]))) vowels.Add(i);
            if (vowels.Count <= 1) return true;

            int last = vowels[vowels.Count - 1];
            bool consonantAfter = last < baseWord.Length - 1;
            int correctPos = tonePos;

            if (vowels.Count == 2)
            {
                char v0 = char.ToLowerInvariant(baseWord[vowels[0]]);
                char v1 = char.ToLowerInvariant(baseWord[vowels[1]]);
                if (consonantAfter) correctPos = vowels[1];
                else
                {
                    // Kiểu mới (thẩm mỹ): ia/ua/ưa, oa/oe, uy đặt dấu ở nguyên âm đầu.
                    bool firstVowel =
                        (v0 == 'i' && v1 == 'a') || (v0 == 'u' && v1 == 'a') || (v0 == 'ư' && v1 == 'a') ||
                        (v0 == 'o' && v1 == 'a') || (v0 == 'o' && v1 == 'e') || (v0 == 'u' && v1 == 'y');
                    if (!firstVowel) return true;
                    correctPos = vowels[0];
                }
            }
            else if (vowels.Count == 3)
            {
                correctPos = vowels[1];
            }

            if (tonePos != correctPos)
            {
                correctedWord = ApplyTone(baseWord, tone, correctPos);
                return false;
            }
            return true;
        }

        /// <summary>
        /// Trả về biến thể vị trí dấu kiểu cũ của từ (ví dụ "hòa" → "hoà"), null nếu không đổi được.
        /// Dùng để so khớp văn bản gõ theo chuẩn cũ.
        /// </summary>
        public static string GetAlternativeTonePlacement(string word)
        {
            if (string.IsNullOrEmpty(word)) return null;
            var lower = word.ToLowerInvariant();
            foreach (var kv in ToneVariantMap)
            {
                int idx = lower.IndexOf(kv.Key, System.StringComparison.Ordinal);
                if (idx < 0) continue;
                if (idx > 0 && lower[idx - 1] == 'q') continue; // "quá" không áp dụng
                var replaced = lower.Substring(0, idx) + kv.Value + lower.Substring(idx + kv.Key.Length);
                if (char.IsUpper(word[0]))
                {
                    if (word.Length > 1 && char.IsUpper(word[1])) return replaced.ToUpperInvariant();
                    return char.ToUpperInvariant(replaced[0]) + (replaced.Length > 1 ? replaced.Substring(1) : "");
                }
                return replaced;
            }
            return null;
        }

        /// <summary>Trả về chuỗi với mọi âm tiết đặt dấu lại đúng theo QĐ 1989.</summary>
        public static string NormalizeAll(string text)
        {
            if (string.IsNullOrEmpty(text)) return text ?? "";
            text = text.Normalize(NormalizationForm.FormC);
            return Regex.Replace(text, @"[\p{L}\p{M}]+", m =>
            {
                IsTonePlacedProperly(m.Value, out var fixedWord);
                return fixedWord;
            });
        }
    }
}
