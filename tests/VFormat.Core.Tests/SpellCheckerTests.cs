using System.Linq;
using VFormat.Core;
using Xunit;

public class SpellCheckerTests
{
    static readonly SpellChecker Sc = new SpellChecker();

    [Theory]
    [InlineData("sát nhập hai xã", "sáp nhập")]
    [InlineData("đã chuẩn đoán bệnh", "chẩn đoán")]
    [InlineData("vô hình chung làm hỏng", "vô hình trung")]
    [InlineData("các cán bộ bổ xung tài liệu", "bổ sung")]
    [InlineData("đề suất tăng ngân sách", "đề xuất")]
    public void CatchesExpandedTypoDictionary(string input, string expected)
    {
        var hits = Sc.Check(input);
        Assert.NotEmpty(hits);
        Assert.Contains(hits, h => h.Suggestion.Contains(expected));
    }

    [Fact]
    public void UpperCaseInputPreservesCase()
    {
        var hits = Sc.Check("SÁT NHẬP TỈNH");
        Assert.NotEmpty(hits);
        Assert.Equal("SÁP NHẬP", hits[0].Suggestion);
    }

    [Fact]
    public void DictionaryHasAtLeastExpectedSize()
    {
        // Hơn 380 cặp lỗi (403 trừ các luật "mở rộng cụm" nguy hiểm) + 50 cặp thiếu dấu.
        Assert.True(SpellCheckData.TypoPairs.GetLength(0) >= 380);
        Assert.True(SpellCheckData.MissingTonePairs.GetLength(0) >= 50);
    }

    [Fact]
    public void CatchesOldTonePlacement()
    {
        var hits = Sc.Check("kế hoạch hoà giải và thuỷ sản");
        Assert.Contains(hits, h => h.Found == "hoà" && h.Suggestion == "hòa");
        Assert.Contains(hits, h => h.Found == "thuỷ" && h.Suggestion == "thủy");
    }

    [Fact]
    public void ConvertsToneMarksBothWays()
    {
        string text = "hoà giải thuỷ sản";
        string newStyle = SpellChecker.ConvertToneMarks(text, newStyle: true);
        Assert.Equal("hòa giải thủy sản", newStyle);

        string oldStyle = SpellChecker.ConvertToneMarks(newStyle, newStyle: false);
        Assert.Equal("hoà giải thuỷ sản", oldStyle);
    }

    [Fact]
    public void ToneCheckDoesNotFlagCorrectWords()
    {
        var text = "Đồng chí Nguyễn Văn Quý, huyện Thụy Anh; kết quả khuyết tật; quả quyết; hoàng hôn; loại; khuỷu tay; thuở nhỏ; người.";
        Assert.DoesNotContain(Sc.Check(text), h => h.Reason == "Sai vị trí đặt dấu thanh");
        Assert.Equal(text, SpellChecker.ConvertToneMarks(text, newStyle: true));
    }

    [Fact]
    public void ToneHitsGiveWordOffsets()
    {
        var text = "Ban hòa giải, Tòa án";
        var hits = SpellChecker.ToneHits(text, newStyle: false);
        Assert.Equal(new[] { "hòa", "Tòa" }, hits.Select(h => h.Found));
        Assert.Equal(new[] { "hoà", "Toà" }, hits.Select(h => h.Suggestion));
        Assert.All(hits, h => Assert.Equal(h.Found, text.Substring(h.Start, h.Length)));
    }

    [Fact]
    public void SyllableValidatorRecognizesVietnameseSyllables()
    {
        Assert.Equal(SyllableValidationResult.ValidSyllable, SyllableValidator.Validate("nghiêng"));
        Assert.Equal(SyllableValidationResult.ValidSyllable, SyllableValidator.Validate("thủy"));
        Assert.Equal(SyllableValidationResult.ForeignOrUnknown, SyllableValidator.Validate("wifi"));
        Assert.Equal(SyllableValidationResult.InvalidSyllable, SyllableValidator.Validate("bcd"));
    }
}
