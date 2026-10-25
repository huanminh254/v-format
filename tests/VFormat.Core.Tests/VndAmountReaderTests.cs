using VFormat.Core;
using Xunit;

public class VndAmountReaderTests
{
    [Theory]
    [InlineData("0", "Không đồng")]
    [InlineData("1", "Một đồng")]
    [InlineData("15", "Mười lăm đồng")]
    [InlineData("21", "Hai mươi mốt đồng")]
    [InlineData("101", "Một trăm linh một đồng")]
    [InlineData("1000", "Một nghìn đồng")]
    [InlineData("1000000", "Một triệu đồng")]
    [InlineData("2345000", "Hai triệu ba trăm bốn mươi lăm nghìn đồng")]
    [InlineData("1000000000", "Một tỷ đồng")]
    public void ConvertsPlainDigits(string digits, string expected)
    {
        var words = VndAmountReader.ToWords(digits, negative: false);
        Assert.Equal(expected, words);
    }

    [Theory]
    [InlineData("2.345.000 đồng")]
    [InlineData("2,345,000 đồng")]
    [InlineData("2 345 000")]
    [InlineData("2345000 VNĐ")]
    public void ParsesGroupedFormats(string input)
    {
        Assert.True(VndAmountReader.TryConvert(input, out var words, out _));
        Assert.Equal("Hai triệu ba trăm bốn mươi lăm nghìn đồng", words);
    }

    [Fact]
    public void RejectsNonInteger()
    {
        Assert.False(VndAmountReader.TryConvert("2.345,50 đồng", out _, out var err));
        Assert.Contains("phần lẻ", err);
    }

    [Fact]
    public void RejectsNonNumeric()
    {
        Assert.False(VndAmountReader.TryConvert("một triệu", out _, out var err));
        Assert.Contains("không phải", err);
    }
}
