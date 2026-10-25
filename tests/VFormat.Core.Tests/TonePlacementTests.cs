using VFormat.Core;
using Xunit;

public class TonePlacementTests
{
    [Theory]
    [InlineData("hòa", true, "hòa")]      // kiểu mới - đúng
    [InlineData("hoà", false, "hòa")]     // kiểu cũ - sửa thành kiểu mới
    [InlineData("thuý", false, "thúy")]
    [InlineData("thủy", true, "thủy")]
    [InlineData("hoàn", true, "hoàn")]    // có phụ âm cuối - dấu trên nguyên âm đầu khi là 2 ngang
    [InlineData("khuya", true, "khuya")]
    public void IsTonePlacedProperly_Works(string word, bool expectedOk, string expectedCorrected)
    {
        var ok = TonePlacementValidator.IsTonePlacedProperly(word, out var corrected);
        Assert.Equal(expectedOk, ok);
        Assert.Equal(expectedCorrected, corrected);
    }

    [Theory]
    [InlineData("hoà bình", "hòa bình")]
    [InlineData("thuý mùa thu", "thúy mùa thu")]
    [InlineData("Thuỷ tinh", "Thủy tinh")]
    public void NormalizeAll_FixesSentence(string input, string expected)
    {
        Assert.Equal(expected, TonePlacementValidator.NormalizeAll(input));
    }

    [Fact]
    public void GetAlternativeTonePlacement_ReturnsOldStyle()
    {
        Assert.Equal("hoà", TonePlacementValidator.GetAlternativeTonePlacement("hòa"));
        Assert.Equal("thuý", TonePlacementValidator.GetAlternativeTonePlacement("thúy"));
        Assert.Null(TonePlacementValidator.GetAlternativeTonePlacement("quốc"));   // 'q' + úa → bỏ qua
    }
}
