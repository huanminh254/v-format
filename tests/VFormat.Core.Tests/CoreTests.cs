using System.Collections.Generic;
using System.Linq;
using VFormat.Core;
using Xunit;

public class CoreTests
{
    static List<ParagraphInfo> Doc(params string[] lines) =>
        lines.Select((t, i) => new ParagraphInfo(t) { Index = i }).ToList();

    static readonly string[] QuyetDinh =
    {
        "ỦY BAN NHÂN DÂN TỈNH HÀ NAM", "SỞ NỘI VỤ", "Số :12 / QĐ- SNV",
        "CỘNG HOÀ XÃ HỘI CHỦ NGHĨA VIỆT NAM", "Độc lập-Tự do-Hạnh phúc", "Hà Nam , ngày 5 tháng 1 năm 2026.",
        "", "QUYẾT ĐỊNH", "Về việc ban hành Quy chế làm việc", "GIÁM ĐỐC SỞ NỘI VỤ",
        "Căn cứ Luật Tổ chức chính quyền địa phương;", "QUYẾT ĐỊNH:",
        "Điều 1 : Ban hành kèm theo Quyết định này Quy chế làm việc.", "Điều 2. Quyết định có hiệu lực kể từ ngày ký.",
        "Nơi nhận:", "+ Như Điều 3", "- UBND tỉnh;", "- Lưu: VT",
        "KT. GIÁM ĐỐC", "PHÓ GIÁM ĐỐC", "", "", "Trần Thị Bình",
    };

    [Fact]
    public void ClassifiesAdministrativeDecision()
    {
        var d = Classifier.Classify(Doc(QuyetDinh), StandardKind.HanhChinh_ND30);
        var k = d.Kinds;
        Assert.Equal(ElementKind.CoQuanChuQuan, k[0]);
        Assert.Equal(ElementKind.CoQuanBanHanh, k[1]);
        Assert.Equal(ElementKind.SoKyHieu, k[2]);
        Assert.Equal(ElementKind.QuocHieu, k[3]);
        Assert.Equal(ElementKind.TieuNgu, k[4]);
        Assert.Equal(ElementKind.DiaDanhNgayThang, k[5]);
        Assert.Equal(ElementKind.TenLoai, k[7]);
        Assert.Equal(ElementKind.TrichYeu, k[8]);
        Assert.Equal(ElementKind.CanCu, k[10]);
        Assert.Equal(ElementKind.Dieu, k[12]);
        Assert.Equal(ElementKind.NoiNhanTitle, k[14]);
        Assert.Equal(ElementKind.NoiNhanItem, k[15]);
        Assert.Equal(ElementKind.NoiNhanItem, k[17]);
        Assert.Equal(ElementKind.ChucVuKy, k[18]);
        Assert.Equal(ElementKind.ChucVuKy, k[19]);
        Assert.Equal(ElementKind.HoTenKy, k[22]);
    }

    [Fact]
    public void FixesTextOfFormalElements()
    {
        var p = Profiles.HanhChinh();
        Assert.Equal("CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM", TextNormalizer.Fix(ElementKind.QuocHieu, QuyetDinh[3], p));
        Assert.Equal("Độc lập - Tự do - Hạnh phúc", TextNormalizer.Fix(ElementKind.TieuNgu, QuyetDinh[4], p));
        Assert.Equal("Số: 12/QĐ-SNV", TextNormalizer.Fix(ElementKind.SoKyHieu, QuyetDinh[2], p));
        Assert.Equal("Số:      /QĐ-SNV", TextNormalizer.Fix(ElementKind.SoKyHieu, "Số: ..../QĐ-SNV", p));
        Assert.Equal("Hà Nam, ngày 05 tháng 01 năm 2026", TextNormalizer.Fix(ElementKind.DiaDanhNgayThang, QuyetDinh[5], p));
        Assert.Equal("Hà Nội, ngày 15 tháng 3 năm 2026", TextNormalizer.FixNgayThang("Hà Nội, ngày 15 tháng 03 năm 2026".Replace("03", "3")));
        Assert.Equal("Điều 1. Ban hành", TextNormalizer.Fix(ElementKind.Dieu, "Điều 1 : Ban hành", p));
        Assert.Equal("- Như Điều 3;", TextNormalizer.Fix(ElementKind.NoiNhanItem, "+ Như Điều 3", p));
        Assert.Equal("- Lưu: VT.", TextNormalizer.Fix(ElementKind.NoiNhanItem, "- Lưu: VT", p, true));
        Assert.Equal("Số 12-NQ/TU", TextNormalizer.Fix(ElementKind.SoKyHieu, "Số: 12 - NQ/TU", Profiles.Dang()));
        Assert.Equal("- Tỉnh ủy (báo cáo),", TextNormalizer.Fix(ElementKind.NoiNhanItem, "- Tỉnh ủy (báo cáo)", Profiles.Dang()));
        Assert.Equal("T/M BAN THƯỜNG VỤ", TextNormalizer.Fix(ElementKind.ChucVuKy, "TM. Ban thường vụ", Profiles.Dang()));
        Assert.Equal("Văn bản (dự thảo), đã gửi.", TextNormalizer.FixSpacing("Văn  bản ( dự thảo ) ,đã gửi ."));
    }

    [Fact]
    public void ChecksFormatAndReportsCodes()
    {
        var paras = Doc(QuyetDinh);
        foreach (var x in paras) { x.FontName = "Times New Roman"; x.FontSize = 14; x.Bold = 0; x.Italic = 0; }
        paras[3].FontName = "Arial";
        var p = Profiles.Doan();
        var d = Classifier.Classify(paras, StandardKind.Doan_HD29);
        var issues = FormatChecker.Check(d, p, new PageInfo { WidthMm = 210, HeightMm = 297, TopMm = 20, BottomMm = 20, LeftMm = 25, RightMm = 15 });
        Assert.Contains(issues, i => i.Code == "DTN-02" && i.Message.Contains("trái"));
        Assert.Contains(issues, i => i.Code == "DTN-03" && i.ParagraphIndex == 3);
        Assert.Contains(issues, i => i.Code == "DTN-07" && i.ParagraphIndex == 4);
        Assert.Contains(issues, i => i.Code == "DTN-08" && i.ParagraphIndex == 5);
        Assert.Contains(issues, i => i.Code == "DTN-10" && i.Message.Contains("ĐOÀN TNCS"));
        Assert.All(issues, i => Assert.Matches(@"^DTN-(0[1-9]|10)$", i.Code));
    }

    [Theory]
    [InlineData("ĐẢNG CỘNG SẢN VIỆT NAM", StandardKind.Dang_HD05)]
    [InlineData("ĐOÀN TNCS HỒ CHÍ MINH", StandardKind.Doan_HD29)]
    [InlineData("BỘ CHỈ HUY QUÂN SỰ TỈNH NAM ĐỊNH", StandardKind.QuanDoi_TT68)]
    [InlineData("KẾ HOẠCH BÀI DẠY", StandardKind.GiaoDuc_CV5512)]
    [InlineData("VĂN BẢN HỢP NHẤT", StandardKind.PhapDien_ND250)]
    [InlineData("ỦY BAN NHÂN DÂN XÃ AN BÌNH", StandardKind.HanhChinh_ND30)]
    public void DetectsStandard(string line, StandardKind expected) =>
        Assert.Equal(expected, StandardDetector.Detect(Doc("", line, "Số: 1/x")));

    [Fact]
    public void SpellCheckerFindsCommonErrors()
    {
        var sc = new SpellChecker(new[] { "# comment", "ko=không" });
        var fixedText = sc.Apply("Sử lý hồ sơ của của uỷ ban nhân dân tỉnh , dần dần ổn định ko trễ");
        Assert.Equal("Xử lý hồ sơ của Ủy ban nhân dân tỉnh, dần dần ổn định không trễ", fixedText);
    }

    [Fact]
    public void AllTemplatesClassifyWithTheirStandard()
    {
        Assert.True(TemplateLibrary.All.Count >= 20);
        foreach (var t in TemplateLibrary.All)
        {
            var paras = Doc(t.Lines.ToArray());
            Assert.Equal(t.Standard, StandardDetector.Detect(paras));
            var d = Classifier.Classify(paras, t.Standard);
            if (t.Standard != StandardKind.GiaoDuc_CV5512)
            {
                Assert.True(d.Has(ElementKind.SoKyHieu), t.Id + " thiếu số ký hiệu");
                Assert.True(d.Has(ElementKind.ChucVuKy), t.Id + " thiếu chức vụ");
                Assert.True(d.Has(ElementKind.HoTenKy), t.Id + " thiếu họ tên");
                Assert.True(d.Has(ElementKind.NoiNhanItem), t.Id + " thiếu nơi nhận");
            }
            else Assert.True(d.Has(ElementKind.TenLoai));
        }
    }

    [Fact]
    public void ClassifiesHeaderAndSignatureInsideTables()
    {
        ParagraphInfo C(string t, int table, int col) => new ParagraphInfo(t) { InTable = true, TableIndex = table, CellColumn = col, CellRow = 1 };
        var paras = new List<ParagraphInfo>
        {
            C("ĐẢNG BỘ TỈNH HÀ NAM\r", 1, 1), C("HUYỆN ỦY KIM BẢNG\r", 1, 1), C("*\r", 1, 1), C("Số 15-NQ/HU\r\a", 1, 1),
            C("ĐẢNG CỘNG SẢN VIỆT NAM\r", 1, 2), C("Kim Bảng, ngày 3 tháng 2 năm 2026\r\a", 1, 2),
            new ParagraphInfo("NGHỊ QUYẾT\r"), new ParagraphInfo("về phát triển kinh tế\r"), new ParagraphInfo("-----\r"),
            new ParagraphInfo("I. TÌNH HÌNH\r"), new ParagraphInfo("Nội dung dài của nghị quyết được trình bày tại đây.\r"),
            C("Nơi nhận:\r", 2, 1), C("- Tỉnh ủy (báo cáo),\r", 2, 1), C("- Lưu Văn phòng Huyện ủy.\r\a", 2, 1),
            C("T/M BAN CHẤP HÀNH\r", 2, 2), C("BÍ THƯ\r", 2, 2), C("\r", 2, 2), C("Nguyễn Văn An\r\a", 2, 2),
        };
        for (int i = 0; i < paras.Count; i++) paras[i].Index = i;
        Assert.Equal(StandardKind.Dang_HD05, StandardDetector.Detect(paras));
        var k = Classifier.Classify(paras, StandardKind.Dang_HD05).Kinds;
        Assert.Equal(new[]
        {
            ElementKind.CoQuanChuQuan, ElementKind.CoQuanBanHanh, ElementKind.DauSao, ElementKind.SoKyHieu,
            ElementKind.TieuDeDang, ElementKind.DiaDanhNgayThang, ElementKind.TenLoai, ElementKind.TrichYeu, ElementKind.DauSao,
            ElementKind.MucLaMa, ElementKind.Body, ElementKind.NoiNhanTitle, ElementKind.NoiNhanItem, ElementKind.NoiNhanItem,
            ElementKind.ChucVuKy, ElementKind.ChucVuKy, ElementKind.Empty, ElementKind.HoTenKy,
        }, k);
        Assert.Equal("Kim Bảng, ngày 03 tháng 02 năm 2026", TextNormalizer.Fix(ElementKind.DiaDanhNgayThang, paras[5].Text, Profiles.Dang()));
    }
}
