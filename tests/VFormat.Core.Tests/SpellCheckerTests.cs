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
}
