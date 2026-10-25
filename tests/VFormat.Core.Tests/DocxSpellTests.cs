using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using VFormat.Core;
using VFormat.Docx;
using Xunit;

public class DocxSpellTests
{
    static readonly SpellChecker Sc = new SpellChecker();

    static Run R(string text, bool bold = false, bool italic = false)
    {
        var r = new Run();
        if (bold || italic)
        {
            var rp = new RunProperties();
            if (bold) rp.AppendChild(new Bold());
            if (italic) rp.AppendChild(new Italic());
            r.AppendChild(rp);
        }
        r.AppendChild(new Text(text) { Space = SpaceProcessingModeValues.Preserve });
        return r;
    }

    static string[] Runs(Paragraph p) =>
        p.Elements<Run>().Select(r => (r.RunProperties?.Bold != null ? "B:" : "") + (r.RunProperties?.Italic != null ? "I:" : "")
            + string.Concat(r.Elements().Select(e => e is TabChar ? "\t" : (e as Text)?.Text))).ToArray();

    [Fact]
    public void KeepsInlineBoldAndItalic()
    {
        var p = new Paragraph(R("Kết quả "), R("hoà", bold: true), R(" nhập, đã "), R("bổ xung", italic: true), R(" hồ sơ."));
        Assert.Equal(2, DocxWriter.FixSpelling(p, Sc));
        Assert.Equal(new[] { "Kết quả ", "B:hòa", " nhập, đã ", "I:bổ sung", " hồ sơ." }, Runs(p));
    }

    [Fact]
    public void SameLengthFixAcrossRunsKeepsEachRunFormat()
    {
        // Word hay cắt run giữa chừng (rsid, kiểm tra chính tả)
        var p = new Paragraph(R("Phòng đã đề "), R("suất", bold: true), R(" phương án, việc sá"), R("t nhập", italic: true));
        Assert.Equal(2, DocxWriter.FixSpelling(p, Sc));
        Assert.Equal(new[] { "Phòng đã đề ", "B:xuất", " phương án, việc sá", "I:p nhập" }, Runs(p));
    }

    [Fact]
    public void ShorterFixAcrossRunsGoesToFirstRun()
    {
        var p = new Paragraph(R("các đơn vị đã "), R("đã", bold: true), R(" thực hiện"));
        Assert.Equal(1, DocxWriter.FixSpelling(p, Sc));
        Assert.Equal("các đơn vị đã thực hiện", DocxModel.TextOf(p));
        Assert.Equal(new[] { "các đơn vị đã", "B:", " thực hiện" }, Runs(p));
    }

    [Fact]
    public void KeepsTabsAndCorrectWords()
    {
        var p = new Paragraph(new Run(new Text("Nguyễn Văn Quý"), new TabChar(), new Text("huyện Thụy Anh")));
        Assert.Equal(0, DocxWriter.FixSpelling(p, Sc));
        Assert.Equal("Nguyễn Văn Quý\thuyện Thụy Anh", DocxModel.TextOf(p));
        Assert.Single(p.Descendants<TabChar>());
    }

    [Fact]
    public void FixesErrorAfterTab()
    {
        var p = new Paragraph(new Run(new Text("Mục 1"), new TabChar(), new Text("sử lý hồ sơ")));
        Assert.Equal(1, DocxWriter.FixSpelling(p, Sc));
        Assert.Equal("Mục 1\txử lý hồ sơ", DocxModel.TextOf(p));
    }

    [Fact]
    public void DecomposedTextIsHandled()
    {
        var p = new Paragraph(R("Ban hoà giải".Normalize(System.Text.NormalizationForm.FormD)));
        Assert.Equal(1, DocxWriter.FixSpelling(p, Sc));
        Assert.Equal("Ban hòa giải", DocxModel.TextOf(p));
    }
}
