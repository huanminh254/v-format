using System;
using System.Collections.Generic;

namespace VFormat.Core
{
    public enum SyllableValidationResult
    {
        ValidSyllable,
        InvalidSyllable,
        ForeignOrUnknown
    }

    /// <summary>
    /// Thẩm định cấu trúc âm tiết tiếng Việt hợp lệ theo quy tắc ngữ âm học
    /// (Phụ âm đầu + Âm đệm/chính + Âm cuối).
    /// </summary>
    public static class SyllableValidator
    {
        private static readonly HashSet<char> Vowels = new HashSet<char>("aàáảãạăằắẳẵặâầấẩẫậeèéẻẽẹêềếểễệiìíỉĩịoòóỏõọôồốổỗộơờớởỡợuùúủũụưừứửữựyỳýỷỹỵ");

        private static readonly string[] ValidOnsets = new string[]
        {
            "ngh", "ng", "nh", "ch", "gh", "gi", "kh", "ph", "qu", "th",
            "tr", "b", "c", "d", "đ", "g", "h", "k", "l", "m",
            "n", "p", "r", "s", "t", "v", "x", ""
        };

        private static readonly string[] ValidCodas = new string[]
        {
            "ng", "nh", "ch", "c", "m", "n", "p", "t", ""
        };

        public static SyllableValidationResult Validate(string word)
        {
            if (string.IsNullOrEmpty(word)) return SyllableValidationResult.InvalidSyllable;

            string text = word.ToLowerInvariant();

            // Ký tự chữ cái quốc tế không thuộc bảng chữ cái tiếng Việt cơ bản
            foreach (char c in text)
            {
                if (c == 'f' || c == 'j' || c == 'w' || c == 'z')
                {
                    return SyllableValidationResult.ForeignOrUnknown;
                }
            }

            // Phải chứa ít nhất 1 nguyên âm tiếng Việt
            bool hasVowel = false;
            for (int i = 0; i < text.Length; i++)
            {
                if (Vowels.Contains(text[i]))
                {
                    hasVowel = true;
                    break;
                }
            }
            if (!hasVowel) return SyllableValidationResult.InvalidSyllable;

            // Tách phụ âm đầu (Onset)
            string onset = null;
            for (int i = 0; i < ValidOnsets.Length; i++)
            {
                string candidate = ValidOnsets[i];
                if (candidate.Length > 0 && text.StartsWith(candidate, StringComparison.Ordinal))
                {
                    string remainder = text.Substring(candidate.Length);
                    bool remainderHasVowel = false;
                    for (int j = 0; j < remainder.Length; j++)
                    {
                        if (Vowels.Contains(remainder[j]))
                        {
                            remainderHasVowel = true;
                            break;
                        }
                    }
                    if (remainderHasVowel)
                    {
                        onset = candidate;
                        break;
                    }
                }
            }
            if (onset == null) onset = "";

            string body = text.Substring(onset.Length);

            // Tách âm cuối (Coda)
            string coda = "";
            for (int i = 0; i < ValidCodas.Length; i++)
            {
                string candidate = ValidCodas[i];
                if (candidate.Length > 0 && body.EndsWith(candidate, StringComparison.Ordinal))
                {
                    coda = candidate;
                    break;
                }
            }

            // Phần nguyên âm hạt nhân (Nucleus)
            string nucleus = coda.Length > 0 ? body.Substring(0, body.Length - coda.Length) : body;
            if (nucleus.Length == 0 || nucleus.Length > 3)
            {
                return SyllableValidationResult.InvalidSyllable;
            }

            for (int i = 0; i < nucleus.Length; i++)
            {
                if (!Vowels.Contains(nucleus[i]))
                {
                    return SyllableValidationResult.InvalidSyllable;
                }
            }

            return SyllableValidationResult.ValidSyllable;
        }
    }
}
