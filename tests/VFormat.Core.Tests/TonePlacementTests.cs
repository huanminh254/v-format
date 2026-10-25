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

    // Từ đúng không được bị "sửa" (lỗi thấy khi chạy thử trên Word 25/10/2026)
    [Theory]
    [InlineData("huyện")] [InlineData("Nguyễn")] [InlineData("khuyết")] [InlineData("quyết")] [InlineData("Quyền")]
    [InlineData("quả")] [InlineData("Quý")] [InlineData("quốc")] [InlineData("quạ")] [InlineData("quỳnh")]
    [InlineData("giữa")] [InlineData("giếng")] [InlineData("gì")] [InlineData("gìn")] [InlineData("giường")]
    [InlineData("người")] [InlineData("mười")] [InlineData("thuở")] [InlineData("hươu")] [InlineData("cửa")] [InlineData("cứu")]
    [InlineData("hoàng")] [InlineData("toán")] [InlineData("loại")] [InlineData("khuỷu")] [InlineData("ngoảnh")] [InlineData("huýt")]
    [InlineData("tài")] [InlineData("mùa")] [InlineData("kìa")] [InlineData("túi")] [InlineData("muốn")] [InlineData("ấy")]
    [InlineData("SẢN")] [InlineData("XUẤT")] [InlineData("NGUYỄN")]
    public void CorrectWordsStayUnchanged(string word)
    {
        Assert.True(TonePlacementValidator.IsTonePlacedProperly(word, out var corrected), $"{word} → {corrected}");
        Assert.Equal(word, TonePlacementValidator.PlaceTone(word, newStyle: false));
    }

    [Theory]
    [InlineData("huỵên", "huyện")] [InlineData("qủa", "quả")] [InlineData("khuýêt", "khuyết")]
    [InlineData("ngừơi", "người")] [InlineData("lọai", "loại")] [InlineData("hòang", "hoàng")] [InlineData("khủyu", "khuỷu")]
    [InlineData("Thuỵ", "Thụy")] [InlineData("THUỶ", "THỦY")] [InlineData("khoẻ", "khỏe")]
    public void MisplacedToneIsMoved(string word, string expected)
    {
        Assert.False(TonePlacementValidator.IsTonePlacedProperly(word, out var corrected));
        Assert.Equal(expected, corrected);
    }

    [Theory]
    [InlineData("thựchiện")] [InlineData("Vieejt")] [InlineData("café")] [InlineData("Đà")] [InlineData("abc")]
    public void NonSyllablesAreLeftAlone(string word)
    {
        Assert.Equal(word, TonePlacementValidator.PlaceTone(word, true));
        Assert.Equal(word, TonePlacementValidator.PlaceTone(word, false));
    }

    [Fact]
    public void OldStyleOnlyMovesOaOeUyWithoutCoda()
    {
        Assert.Equal("Ban hoà giải, Thuỷ sản, Toà án huyện, khoẻ, loại, hoàng hôn, Nguyễn Văn Quý",
            TonePlacementValidator.NormalizeAll("Ban hòa giải, Thủy sản, Tòa án huyện, khỏe, loại, hoàng hôn, Nguyễn Văn Quý", newStyle: false));
        // đã ở kiểu cũ thì giữ nguyên (không đảo chiều)
        Assert.Equal("hoà", TonePlacementValidator.PlaceTone("hoà", newStyle: false));
        Assert.Equal("hòa", TonePlacementValidator.GetAlternativeTonePlacement("hoà"));
    }
}
