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
    public void TheoDeNghiAfterCanCuIsCanCu()
    {
        var lines = QuyetDinh.ToList();
        lines.Insert(11, "Theo đề nghị của Chánh Văn phòng Sở.");
        var d = Classifier.Classify(Doc(lines.ToArray()), StandardKind.HanhChinh_ND30);
        Assert.Equal(ElementKind.CanCu, d.Kinds[11]);
        Assert.True(Planner.PlanParagraph(ElementKind.CanCu, d.Paragraphs[11], Profiles.HanhChinh()).Italic);
    }

    [Theory]
    [InlineData("Điều 1. Ban hành kèm theo Quyết định này Quy chế làm việc của Sở Nội vụ.", 7)]
    [InlineData("Điều 3. Các cơ quan có liên quan chịu trách nhiệm thi hành Quyết định này./.", 7)]
    [InlineData("Điều 5. Phạm vi điều chỉnh", 0)]
    public void DieuWithContentBoldsOnlyPrefix(string text, int boldPrefix)
    {
        var plan = Planner.PlanParagraph(ElementKind.Dieu, new ParagraphInfo(text), Profiles.HanhChinh());
        Assert.Equal(boldPrefix, plan.BoldPrefixLength);
        Assert.Equal(boldPrefix == 0, plan.Bold);
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
        Assert.Equal("- Tỉnh ủy (báo cáo);", TextNormalizer.Fix(ElementKind.NoiNhanItem, "- Tỉnh ủy (báo cáo)", Profiles.Dang()));
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
    [InlineData("CÔNG AN TỈNH HÀ NAM", StandardKind.CongAn_BCA08)]
    [InlineData("BỘ CÔNG AN", StandardKind.CongAn_BCA08)]
    [InlineData("BỘ TƯ LỆNH CẢNH SÁT BIỂN", StandardKind.QuanDoi_TT68)]
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

    [Fact]
    public void KeepsSpaceBeforeEllipsisPlaceholder()
    {
        Assert.Equal("Kính gửi: ...", TextNormalizer.FixSpacing("Kính gửi: ..."));
        Assert.Equal("Họ và tên: ...; lớp: ...", TextNormalizer.FixSpacing("Họ và tên: ... ; lớp: ..."));
        Assert.Equal("Hết.", TextNormalizer.FixSpacing("Hết ."));
        Assert.Equal("- ...;", TextNormalizer.FixNoiNhanItem("- ...;", false));
        Assert.Equal("- Lưu: VT, ....", TextNormalizer.FixNoiNhanItem("- Lưu: VT, ....", true));
        Assert.Equal("- Lưu: VT.", TextNormalizer.FixNoiNhanItem("- Lưu: VT.", true));
    }

    [Fact]
    public void ConsolidatedDocumentHasTwoTitleLines()
    {
        var k = Classifier.Classify(Doc("BỘ TÀI CHÍNH", "Số: 1/VBHN-BTC", "CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM",
            "Độc lập - Tự do - Hạnh phúc", "Hà Nội, ngày 01 tháng 3 năm 2026", "VĂN BẢN HỢP NHẤT", "THÔNG TƯ", "Quy định về kế toán",
            "Chương I", "QUY ĐỊNH CHUNG"), StandardKind.PhapDien_ND250).Kinds;
        Assert.Equal(new[] { ElementKind.TenLoai, ElementKind.TenLoai, ElementKind.TrichYeu, ElementKind.Chuong, ElementKind.TenPhanChuong },
            k.Skip(5).ToArray());
    }

    [Fact]
    public void PoliceSignerNeedsRank()
    {
        var lines = QuyetDinh.Select(x => x.Replace("SỞ NỘI VỤ", "PHÒNG THAM MƯU").Replace("ỦY BAN NHÂN DÂN TỈNH HÀ NAM", "CÔNG AN TỈNH HÀ NAM")).ToArray();
        var paras = Doc(lines);
        Assert.Equal(StandardKind.CongAn_BCA08, StandardDetector.Detect(paras));
        var p = Profiles.CongAn();
        var issues = FormatChecker.Check(Classifier.Classify(paras, p.Kind), p);
        Assert.Contains(issues, i => i.Code == "CA08-09" && i.Message.Contains("cấp bậc hàm") && !i.AutoFixable);

        lines[lines.Length - 1] = "Đại tá Trần Thị Bình";
        var cd = Classifier.Classify(Doc(lines), p.Kind);
        Assert.Equal(ElementKind.HoTenKy, cd.Kinds.Last());
        Assert.DoesNotContain(FormatChecker.Check(cd, p), i => i.Message.Contains("cấp bậc hàm"));
    }

    [Fact]
    public void MilitarySignerNeedsRank()
    {
        var lines = QuyetDinh.Select(x => x.Replace("SỞ NỘI VỤ", "PHÒNG THAM MƯU").Replace("ỦY BAN NHÂN DÂN TỈNH HÀ NAM", "BỘ CHỈ HUY QUÂN SỰ TỈNH HÀ NAM")).ToArray();
        var paras = Doc(lines);
        Assert.Equal(StandardKind.QuanDoi_TT68, StandardDetector.Detect(paras));
        var p = Profiles.QuanDoi();
        Assert.Contains(FormatChecker.Check(Classifier.Classify(paras, p.Kind), p),
            i => i.Code == "QD68-09" && i.Message.Contains("cấp bậc hàm") && !i.AutoFixable);

        lines[lines.Length - 1] = "Thượng tá Trần Thị Bình";
        Assert.DoesNotContain(FormatChecker.Check(Classifier.Classify(Doc(lines), p.Kind), p), i => i.Message.Contains("cấp bậc hàm"));
    }

    static readonly string[] CongVanDang =
    {
        "ĐẢNG BỘ TỈNH HÀ NAM", "HUYỆN ỦY KIM BẢNG", "*", "Số 12-CV/HU", "V/v triển khai học tập nghị quyết",
        "ĐẢNG CỘNG SẢN VIỆT NAM", "Kim Bảng, ngày 05 tháng 3 năm 2026",
        "Kính gửi: Các chi bộ, đảng bộ cơ sở.",
        "Thực hiện kế hoạch của Tỉnh ủy, Ban Thường vụ Huyện ủy đề nghị các chi bộ triển khai học tập nghị quyết.",
        "Nơi nhận:", "- Như trên,", "- Lưu Văn phòng.",
        "T/M BAN THƯỜNG VỤ", "PHÓ BÍ THƯ", "", "Trần Văn Bình",
    };

    [Fact]
    public void PartyOfficialLetterSubjectIsItalicSize12()
    {
        var p = Profiles.Dang();
        var paras = Doc(CongVanDang);
        var cd = Classifier.Classify(paras, p.Kind);
        int i = cd.Kinds.IndexOf(ElementKind.TrichYeuCongVan);
        Assert.Equal(4, i);
        var plan = Planner.PlanParagraph(ElementKind.TrichYeuCongVan, paras[i], p);
        Assert.Equal(12f, plan.Size);
        Assert.True(plan.Italic);

        paras[i].FontSize = 13; paras[i].Italic = 0;
        var issues = FormatChecker.Check(cd, p).Where(x => x.ParagraphIndex == i).ToList();
        Assert.Contains(issues, x => x.Code == "D05-04" && x.Message.Contains("cỡ 12"));
        Assert.Contains(issues, x => x.Code == "D05-05" && x.Message.Contains("in nghiêng"));
    }

    [Theory]
    [InlineData(13f, 14f)]
    [InlineData(14f, 14f)]
    public void PartyBodySizeStaysWithinRange(float chosen, float expected)
    {
        var p = Profiles.Dang().Clone();
        p.SetBodySize(chosen);
        Assert.Equal(expected, p.BodySize);
        foreach (var k in FormatProfile.BodySizedKinds) Assert.Equal(expected, p.Styles[k].Size);

        var hc = Profiles.HanhChinh().Clone();
        hc.SetBodySize(13);
        Assert.Equal(13f, hc.BodySize);
        Assert.Equal(13f, hc.Styles[ElementKind.Dieu].Size);
    }

    [Theory]
    [InlineData(13f, true)]
    [InlineData(14f, false)]
    [InlineData(15f, false)]
    public void PartyContentElementsAre14To15(float size, bool flagged)
    {
        var p = Profiles.Dang();
        foreach (var k in FormatProfile.BodySizedKinds)
        {
            Assert.Equal(14f, p.Styles[k].MinSize);
            Assert.Equal(15f, p.Styles[k].MaxSize);
        }
        var paras = Doc(CongVanDang);
        var cd = Classifier.Classify(paras, p.Kind);
        int kinhGui = cd.Kinds.IndexOf(ElementKind.KinhGui), body = cd.Kinds.IndexOf(ElementKind.Body);
        Assert.True(kinhGui >= 0 && body >= 0);
        paras[kinhGui].FontSize = size; paras[body].FontSize = size;
        var sizeIssues = FormatChecker.Check(cd, p).Where(x => x.Code == "D05-04").Select(x => x.ParagraphIndex).ToList();
        Assert.Equal(flagged, sizeIssues.Contains(kinhGui));
        Assert.Equal(flagged, sizeIssues.Contains(body));
    }

    // ===== NĐ 30: số văn bản, cỡ chữ khối đầu =====

    [Theory]
    [InlineData("Số: 5/QĐ-UBND", "Số: 05/QĐ-UBND")]
    [InlineData("Số : 7 / CV - VP", "Số: 07/CV-VP")]
    [InlineData("Số: 12/QĐ-UBND", "Số: 12/QĐ-UBND")]
    [InlineData("Số: 1/2026/VBHN-LQ-VPQH", "Số: 01/2026/VBHN-LQ-VPQH")]
    [InlineData("Số: ..../QĐ-UBND", "Số:      /QĐ-UBND")]
    public void DocumentNumberBelowTenGetsLeadingZero(string input, string expected)
    {
        Assert.Equal(expected, TextNormalizer.Fix(ElementKind.SoKyHieu, input, Profiles.HanhChinh()));
        Assert.Equal("Số 05-NQ/TU", TextNormalizer.Fix(ElementKind.SoKyHieu, "Số 5-NQ/TU", Profiles.Dang()));

        var p = Profiles.HanhChinh();
        var lines = QuyetDinh.ToArray(); lines[2] = "Số: 5/QĐ-SNV";
        Assert.Contains(FormatChecker.Check(Classifier.Classify(Doc(lines), p.Kind), p), i => i.Code == "ND30-08" && i.Message.Contains("05/QĐ-SNV"));
    }

    [Theory]
    [InlineData(12f, 13f, 13f, 0)]
    [InlineData(13f, 14f, 14f, 0)]
    [InlineData(12f, 13f, 14f, 1)]
    [InlineData(13f, 13f, 13f, 1)]
    public void HeaderFontSizesMustBeConsistent(float quocHieu, float tieuNgu, float ngayThang, int expected)
    {
        var p = Profiles.HanhChinh();
        var paras = Doc(QuyetDinh);
        paras[3].FontSize = quocHieu; paras[4].FontSize = tieuNgu; paras[5].FontSize = ngayThang;
        var issues = FormatChecker.Check(Classifier.Classify(paras, p.Kind), p).Where(i => i.Message.Contains("chưa thống nhất")).ToList();
        Assert.Equal(expected, issues.Count);
        Assert.All(issues, i => Assert.Equal("ND30-07", i.Code));

        // thông số mặc định tự thống nhất: NĐ 30 12/13/13, Quân đội 13/14/14
        foreach (var (prof, qh) in new[] { (Profiles.HanhChinh(), 12f), (Profiles.CongAn(), 12f), (Profiles.QuanDoi(), 13f) })
        {
            Assert.Equal(qh, prof.Styles[ElementKind.QuocHieu].Size);
            Assert.Equal(qh + 1, prof.Styles[ElementKind.TieuNgu].Size);
            Assert.Equal(qh + 1, prof.Styles[ElementKind.DiaDanhNgayThang].Size);
        }
    }

    // ===== Đảng: HD 05-HD/VPTW ngày 27/5/2026 =====

    static readonly string[] BaoCaoDang =
    {
        "ĐẢNG BỘ TỈNH HÀ NAM", "HUYỆN ỦY KIM BẢNG", "*", "Số 15-BC/HU",
        "ĐẢNG CỘNG SẢN VIỆT NAM", "Kim Bảng, ngày 05 tháng 3 năm 2026",
        "BÁO CÁO", "kết quả đại hội chi bộ nhiệm kỳ 2025 - 2030",
        "Căn cứ Điều lệ Đảng", "- Căn cứ Quy chế làm việc của Huyện ủy",
        "1. Công tác chuẩn bị đại hội được thực hiện đúng kế hoạch",
        "Nơi nhận:", "- Tỉnh ủy,", "- Lưu Văn phòng Huyện ủy.",
        "K/T BÍ THƯ", "PHÓ BÍ THƯ", "", "Trần Văn Bình",
    };

    [Theory]
    [InlineData(20f, 20f, 30f, 15f, false)]
    [InlineData(25f, 20f, 30f, 15f, true)]
    [InlineData(20f, 20f, 35f, 15f, true)]
    [InlineData(20f, 20f, 30f, 20f, true)]
    public void PartyMarginsAreFixed(float top, float bottom, float left, float right, bool flagged)
    {
        var p = Profiles.Dang();
        Assert.Equal((20f, 20f, 30f, 15f), (p.MarginTopMm, p.MarginBottomMm, p.MarginLeftMm, p.MarginRightMm));
        var issues = FormatChecker.Check(Classifier.Classify(Doc(BaoCaoDang), p.Kind), p,
            new PageInfo { WidthMm = 210, HeightMm = 297, TopMm = top, BottomMm = bottom, LeftMm = left, RightMm = right });
        Assert.Equal(flagged, issues.Any(i => i.Code == "D05-02"));
    }

    [Fact]
    public void PartyBodyUsesExactLineSpacing()
    {
        var p = Profiles.Dang().Clone();
        p.SetLineSpacing(1.5f);
        var cd = Classifier.Classify(Doc(BaoCaoDang), p.Kind);
        int body = cd.Kinds.IndexOf(ElementKind.MucSo);
        var plan = Planner.PlanParagraph(cd.Kinds[body], cd.Paragraphs[body], p);
        Assert.Equal(18f, plan.LineExactPt);
        Assert.Null(plan.LineMultiple);
        Assert.Null(Planner.PlanParagraph(ElementKind.TieuDeDang, cd.Paragraphs[4], p).LineExactPt);   // khối đầu: dòng đơn
        Assert.Null(Planner.PlanParagraph(ElementKind.Body, new ParagraphInfo("Nội dung"), Profiles.HanhChinh()).LineExactPt);
    }

    [Theory]
    [InlineData(StandardKind.Dang_HD05, "T/M BAN THƯỜNG VỤ", true)]
    [InlineData(StandardKind.Dang_HD05, "K/T BÍ THƯ", true)]
    [InlineData(StandardKind.Dang_HD05, "PHÓ BÍ THƯ", false)]
    [InlineData(StandardKind.Doan_HD29, "TM. BAN THƯỜNG VỤ", true)]
    [InlineData(StandardKind.Doan_HD29, "BÍ THƯ", false)]
    [InlineData(StandardKind.HanhChinh_ND30, "PHÓ GIÁM ĐỐC", true)]
    public void SignerTitleBoldDependsOnStandard(StandardKind kind, string line, bool bold)
    {
        var p = Profiles.Get(kind);
        Assert.Equal(bold, Planner.PlanParagraph(ElementKind.ChucVuKy, new ParagraphInfo(line), p).Bold);

        var cd = Classifier.Classify(Doc(BaoCaoDang), StandardKind.Dang_HD05);
        int chucVu = cd.Kinds.LastIndexOf(ElementKind.ChucVuKy);
        Assert.Equal("PHÓ BÍ THƯ", cd.Paragraphs[chucVu].Text);
        cd.Paragraphs[chucVu].Bold = 1;
        Assert.Contains(FormatChecker.Check(cd, Profiles.Dang()), i => i.ParagraphIndex == chucVu && i.Message.Contains("không in đậm"));
    }

    [Theory]
    [InlineData(StandardKind.Dang_HD05, 9, false)]
    [InlineData(StandardKind.Doan_HD29, 9, true)]
    [InlineData(StandardKind.HanhChinh_ND30, 0, false)]
    public void KinhGuiWordIsItalic(StandardKind kind, int italicPrefix, bool bold)
    {
        var plan = Planner.PlanParagraph(ElementKind.KinhGui, new ParagraphInfo("Kính gửi: Ban Thường vụ Tỉnh đoàn"), Profiles.Get(kind));
        Assert.Equal(italicPrefix, plan.ItalicPrefixLength);
        Assert.False(plan.Italic);
        Assert.Equal(bold, plan.Bold);
    }

    [Fact]
    public void PartyClausesAndHeadings()
    {
        var p = Profiles.Dang();
        var cd = Classifier.Classify(Doc(BaoCaoDang), p.Kind);
        int khoan = cd.Kinds.IndexOf(ElementKind.MucSo);
        Assert.False(Planner.PlanParagraph(ElementKind.MucSo, cd.Paragraphs[khoan], p).Bold);   // khoản: đứng
        Assert.True(Planner.PlanParagraph(ElementKind.MucSo, cd.Paragraphs[khoan], Profiles.HanhChinh()).Bold);
        foreach (var k in new[] { ElementKind.Phan, ElementKind.Chuong, ElementKind.Muc, ElementKind.TenPhanChuong })
            Assert.Equal((14f, 15f), (p.Styles[k].MinSize, p.Styles[k].MaxSize));
    }

    [Fact]
    public void PartyRecipientsEndWithSemicolon()
    {
        var p = Profiles.Dang();
        var cd = Classifier.Classify(Doc(BaoCaoDang), p.Kind);
        int item = cd.Kinds.IndexOf(ElementKind.NoiNhanItem);
        Assert.Contains(FormatChecker.Check(cd, p), i => i.ParagraphIndex == item && i.Message.Contains("\"- Tỉnh ủy;\""));
        Assert.All(TemplateLibrary.All.Where(t => t.Standard == StandardKind.Dang_HD05)
            .SelectMany(t => t.Lines).Where(l => l.StartsWith("- ") && !l.Contains("Lưu") && !l.Contains("Căn cứ")),
            l => Assert.EndsWith(";", l));
    }

    [Theory]
    [InlineData(StandardKind.Dang_HD05, "-----")]
    [InlineData(StandardKind.Doan_HD29, "-------")]
    [InlineData(StandardKind.HanhChinh_ND30, null)]
    public void TitleDashesUnderSummary(StandardKind kind, string dashes)
    {
        var p = Profiles.Get(kind);
        var cd = Classifier.Classify(Doc(BaoCaoDang), kind);
        int trichYeu = cd.Kinds.IndexOf(ElementKind.TrichYeu);
        Assert.Equal(7, trichYeu);
        Assert.Equal(dashes == null ? -1 : trichYeu, Planner.PlanTitleDashes(cd, p));
        Assert.Equal(dashes != null, FormatChecker.Check(cd, p).Any(i => i.Message.Contains("Thiếu dòng gạch")));
        if (dashes == null) return;

        var lines = BaoCaoDang.ToList(); lines.Insert(8, dashes);
        var fixedDoc = Classifier.Classify(Doc(lines.ToArray()), kind);
        Assert.Equal(ElementKind.DauSao, fixedDoc.Kinds[8]);
        Assert.Equal(-1, Planner.PlanTitleDashes(fixedDoc, p));
        Assert.DoesNotContain(FormatChecker.Check(fixedDoc, p), i => i.Message.Contains("Thiếu dòng gạch"));
    }

    [Theory]
    [InlineData(StandardKind.Dang_HD05, "KT. Chánh văn phòng", "K/T CHÁNH VĂN PHÒNG")]
    [InlineData(StandardKind.Dang_HD05, "TL. Ban thường vụ", "T/L BAN THƯỜNG VỤ")]
    [InlineData(StandardKind.Dang_HD05, "TUQ. Bí thư", "T/UQ BÍ THƯ")]
    [InlineData(StandardKind.HanhChinh_ND30, "K/T GIÁM ĐỐC", "KT. GIÁM ĐỐC")]
    [InlineData(StandardKind.HanhChinh_ND30, "T/L BỘ TRƯỞNG", "TL. BỘ TRƯỞNG")]
    public void SigningAuthorityAbbreviationFollowsStandard(StandardKind kind, string input, string expected) =>
        Assert.Equal(expected, TextNormalizer.Fix(ElementKind.ChucVuKy, input, Profiles.Get(kind)));

    [Fact]
    public void PartyLegalBasesUseDashAndSemicolon()
    {
        var p = Profiles.Dang();
        var cd = Classifier.Classify(Doc(BaoCaoDang), p.Kind);
        Assert.Equal(new[] { ElementKind.CanCu, ElementKind.CanCu }, cd.Kinds.Skip(8).Take(2).ToArray());
        Assert.Equal("- Căn cứ Điều lệ Đảng;", TextNormalizer.Fix(ElementKind.CanCu, "Căn cứ Điều lệ Đảng", p));
        Assert.Equal("- Căn cứ Quy chế làm việc của Huyện ủy.", TextNormalizer.Fix(ElementKind.CanCu, "- Căn cứ Quy chế làm việc của Huyện ủy;", p, true));
        Assert.True(Planner.NeedsTextFix(ElementKind.CanCu, p));
        Assert.False(Planner.PlanParagraph(ElementKind.CanCu, cd.Paragraphs[8], p).Italic);   // HD 05: căn cứ kiểu chữ đứng

        var hc = Profiles.HanhChinh();   // NĐ 30 giữ nguyên căn cứ in nghiêng, không thêm gạch đầu dòng
        Assert.Equal("Căn cứ Luật Tổ chức chính quyền địa phương;", TextNormalizer.Fix(ElementKind.CanCu, "Căn cứ Luật Tổ chức chính quyền địa phương;", hc, true));
        Assert.False(Planner.NeedsTextFix(ElementKind.CanCu, hc));
        Assert.True(Planner.PlanParagraph(ElementKind.CanCu, cd.Paragraphs[8], hc).Italic);
    }

    // ===== Đoàn: HD 29-HD/TWĐTN-VP, Mẫu 3 =====

    [Fact]
    public void YouthUnionPageAndBody()
    {
        var p = Profiles.Doan().Clone();
        Assert.Equal((20f, 20f, 30f, 20f), (p.MarginTopMm, p.MarginBottomMm, p.MarginLeftMm, p.MarginRightMm));
        Assert.Equal((20f, 20f), (p.MinRightMm, p.MaxRightMm));

        p.SetLineSpacing(1.5f);
        Assert.Equal(1.2f, p.LineSpacing);
        p.SetBodySize(13);
        Assert.Equal(14f, p.BodySize);   // phần lời văn cỡ 14 cố định
        var plan = Planner.PlanParagraph(ElementKind.Body, new ParagraphInfo("Nội dung văn bản."), p);
        Assert.Equal(1.27f * 28.3465f, plan.FirstIndentPt.Value, 2);
        Assert.Equal(1.2f, plan.LineMultiple);

        var issues = FormatChecker.Check(Classifier.Classify(Doc(BaoCaoDang), p.Kind), p,
            new PageInfo { WidthMm = 210, HeightMm = 297, TopMm = 20, BottomMm = 20, LeftMm = 30, RightMm = 15 });
        Assert.Contains(issues, i => i.Code == "DTN-02" && i.Message.Contains("Lề phải"));
    }

    [Theory]
    [InlineData(ElementKind.DiaDanhNgayThang, 13f, false, true)]
    [InlineData(ElementKind.TrichYeuCongVan, 12f, false, true)]
    [InlineData(ElementKind.NoiNhanTitle, 13f, true, false)]
    [InlineData(ElementKind.NoiNhanItem, 11f, false, false)]
    [InlineData(ElementKind.Phan, 15f, true, false)]
    public void YouthUnionElementStyles(ElementKind kind, float size, bool bold, bool italic)
    {
        var st = Profiles.Doan().Styles[kind];
        Assert.Equal((size, bold, italic), (st.Size, st.Bold, st.Italic));
        Assert.Equal(size, st.MinSize);
        var plan = Planner.PlanParagraph(kind, new ParagraphInfo(kind == ElementKind.NoiNhanTitle ? "Nơi nhận:" : "x"), Profiles.Doan());
        Assert.Equal(0, plan.UnderlinePrefixLength);   // "Nơi nhận" của Đoàn không gạch chân
    }

    // ===== Lệch phát hiện khi chạy 30 file test/ trên Word =====

    [Fact]
    public void TextColorIsResetToBlack()
    {
        var p = Profiles.HanhChinh();
        foreach (var k in new[] { ElementKind.Body, ElementKind.QuocHieu, ElementKind.HoTenKy, ElementKind.TableText })
            Assert.True(Planner.PlanParagraph(k, new ParagraphInfo("Nội dung") { InTable = k == ElementKind.TableText }, p).ResetColor);

        var paras = Doc(QuyetDinh);
        paras[13].Colored = true;
        var issues = FormatChecker.Check(Classifier.Classify(paras, p.Kind), p);
        Assert.Contains(issues, i => i.ParagraphIndex == 13 && i.Code == "ND30-05" && i.Message.Contains("màu đen"));
        Assert.DoesNotContain(issues, i => i.ParagraphIndex == 12 && i.Message.Contains("màu"));
    }

    [Theory]
    [InlineData(1, 1, false, false, 2)]
    [InlineData(-1, -1, null, null, 0)]   // nhấn mạnh vài chữ trong đoạn: giữ nguyên
    [InlineData(0, 0, null, null, 0)]
    public void BodyLosesWholeParagraphEmphasis(int bold, int italic, bool? planBold, bool? planItalic, int issues)
    {
        var p = Profiles.HanhChinh();
        var info = new ParagraphInfo("Các bộ phận kiểm tra thành phần hồ sơ và gửi kết quả về tổng hợp.") { Bold = bold, Italic = italic };
        var plan = Planner.PlanParagraph(ElementKind.Body, info, p);
        Assert.Equal(planBold, plan.Bold);
        Assert.Equal(planItalic, plan.Italic);

        var paras = Doc(QuyetDinh.Take(13).Concat(new[] { info.Text }).Concat(QuyetDinh.Skip(14)).ToArray());
        paras[13].Bold = bold; paras[13].Italic = italic;
        var cd = Classifier.Classify(paras, p.Kind);
        Assert.Equal(ElementKind.Body, cd.Kinds[13]);
        Assert.Equal(issues, FormatChecker.Check(cd, p).Count(i => i.ParagraphIndex == 13 && i.Message.Contains("cả đoạn")));
    }

    [Theory]
    [InlineData(297f, 210f, false, true)]    // ngang, không bảng nội dung → phải xoay dọc
    [InlineData(297f, 210f, true, false)]    // ngang vì có bảng → hợp lệ
    [InlineData(210f, 297f, false, false)]
    [InlineData(216f, 279f, false, true)]    // Letter
    public void LandscapeOnlyForContentTables(float w, float h, bool contentTable, bool flagged)
    {
        var p = Profiles.HanhChinh();
        var paras = Doc(QuyetDinh);
        if (contentTable) paras.Insert(14, new ParagraphInfo("Số liệu") { InTable = true, TableIndex = 3, CellRow = 1, CellColumn = 1 });
        var cd = Classifier.Classify(paras, p.Kind);
        Assert.Equal(contentTable, Planner.HasContentTable(cd));
        var issues = FormatChecker.Check(cd, p, new PageInfo { WidthMm = w, HeightMm = h, TopMm = 20, BottomMm = 20, LeftMm = 30, RightMm = 15 });
        Assert.Equal(flagged, issues.Any(i => i.Code == "ND30-01"));
    }

    [Theory]
    [InlineData(StandardKind.Dang_HD05, "-----", "*")]
    [InlineData(StandardKind.Dang_HD05, "***", "*")]
    [InlineData(StandardKind.Doan_HD29, "*", "***")]
    [InlineData(StandardKind.Doan_HD29, "***", "***")]
    public void HeaderSeparatorFollowsStandard(StandardKind kind, string input, string expected)
    {
        var p = Profiles.Get(kind);
        var lines = BaoCaoDang.ToArray(); lines[2] = input;
        if (kind == StandardKind.Doan_HD29) lines[4] = "ĐOÀN TNCS HỒ CHÍ MINH";
        var cd = Classifier.Classify(Doc(lines), kind);
        Assert.Equal(2, Planner.HeaderSeparatorIndex(cd));
        Assert.True(Planner.NeedsTextFix(cd, 2, p));
        Assert.Equal(expected, Planner.FixText(cd, 2, p, input));
        Assert.Equal(input != expected, FormatChecker.Check(cd, p).Any(i => i.ParagraphIndex == 2 && i.Code.EndsWith("-07")));

        // dòng gạch dưới trích yếu không bị đổi thành dấu sao
        var withDashes = lines.ToList(); withDashes.Insert(8, p.TitleDashes);
        var cd2 = Classifier.Classify(Doc(withDashes.ToArray()), kind);
        Assert.Equal(p.TitleDashes, Planner.FixText(cd2, 8, p, p.TitleDashes));
        Assert.False(Planner.NeedsTextFix(cd2, 8, p));
    }

    [Fact]
    public void HeaderTableKeepsSingleSeparator()
    {
        var p = Profiles.Dang();
        var lines = BaoCaoDang.ToArray(); lines[2] = "-----";
        var plan = Planner.PlanHeader(Classifier.Classify(Doc(lines), p.Kind), p);
        Assert.Equal(new[] { "ĐẢNG BỘ TỈNH HÀ NAM", "HUYỆN ỦY KIM BẢNG", "-----", "Số 15-BC/HU" }, plan.Left.ToArray());
        Assert.Equal(-1, Planner.HeaderSeparatorIndex(Classifier.Classify(Doc(QuyetDinh), StandardKind.HanhChinh_ND30)));
    }

    // ===== Lệch 5–10 =====

    static List<ParagraphInfo> CongVanCongAn()
    {
        var paras = Doc("CÔNG AN TỈNH HÀ NAM", "PHÒNG THAM MƯU", "Số : 11 / CV - PV01",
            "CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM", "Độc lập - Tự do - Hạnh phúc", "Hà Nam, ngày 24 tháng 10 năm 2026",
            "công văn", "V/v triển khai công tác bảo đảm an ninh trật tự", "Kính gửi: Các đơn vị trực thuộc",
            "Đề nghị các đơn vị nghiêm túc triển khai thực hiện./.");
        foreach (var (i, col) in new[] { (0, 1), (1, 1), (2, 1), (3, 2), (4, 2), (5, 2) })
        { paras[i].InTable = true; paras[i].TableIndex = 1; paras[i].CellRow = 1; paras[i].CellColumn = col; }
        return paras;
    }

    [Fact]
    public void OfficialLetterSubjectMovesUnderNumber()
    {
        var p = Profiles.CongAn();
        var cd = Classifier.Classify(CongVanCongAn(), p.Kind);
        Assert.Equal(6, Planner.StrayCongVanLabel(cd));
        Assert.Equal((7, 2), Planner.PlanCongVanSubjectMove(cd));
        var issues = FormatChecker.Check(cd, p);
        Assert.Contains(issues, i => i.ParagraphIndex == 6 && i.Message.Contains("không ghi tên loại"));
        Assert.Contains(issues, i => i.ParagraphIndex == 7 && i.Message.Contains("dưới số, ký hiệu"));

        // sau khi chuyển: trích yếu trong ô trái dưới số, không còn nhãn
        var moved = CongVanCongAn();
        moved.RemoveAt(7); moved.RemoveAt(6);
        moved.Insert(3, new ParagraphInfo("V/v triển khai công tác bảo đảm an ninh trật tự") { InTable = true, TableIndex = 1, CellRow = 1, CellColumn = 1 });
        var cd2 = Classifier.Classify(moved, p.Kind);
        Assert.Equal(ElementKind.TrichYeuCongVan, cd2.Kinds[3]);
        Assert.Equal((-1, -1), Planner.PlanCongVanSubjectMove(cd2));
        Assert.Equal(-1, Planner.StrayCongVanLabel(cd2));
    }

    [Fact]
    public void TableCellsLoseWholeCellEmphasisAndGetContentWidths()
    {
        var plan = Planner.PlanParagraph(ElementKind.TableText, new ParagraphInfo("Câu trả lời ban đầu.") { InTable = true, Bold = 1, Italic = 1 }, Profiles.GiaoDuc());
        Assert.Equal((false, false, true), (plan.Bold, plan.Italic, plan.ClearUnderline));

        var paras = new List<ParagraphInfo>();
        void Row(int r, string a, string b)
        {
            paras.Add(new ParagraphInfo(a) { Index = paras.Count, InTable = true, TableIndex = 1, CellRow = r, CellColumn = 1, Bold = 1 });
            paras.Add(new ParagraphInfo(b) { Index = paras.Count, InTable = true, TableIndex = 1, CellRow = r, CellColumn = 2, Bold = 1 });
        }
        Row(1, "Hoạt động của giáo viên và học sinh", "Sản phẩm dự kiến");
        Row(2, "Hoạt động 1 Mở đầu. Giáo viên nêu tình huống và yêu cầu học sinh dự đoán.", "Câu trả lời ban đầu.");
        var cd = Classifier.Classify(paras, StandardKind.GiaoDuc_CV5512);
        var w = Planner.PlanColumnWidths(cd, 1, 160);
        Assert.Equal(160f, w.Sum(), 1);
        Assert.True(w[0] > 2 * w[1]);                    // cột chữ dài rộng hơn
        Assert.All(w, x => Assert.True(x >= 12f));

        var issues = FormatChecker.Check(cd, Profiles.GiaoDuc());
        Assert.DoesNotContain(issues, i => i.ParagraphIndex <= 1 && i.Message.Contains("cả ô"));   // hàng tiêu đề được đậm
        Assert.Contains(issues, i => i.ParagraphIndex == 2 && i.Message.Contains("cả ô"));
        Assert.Equal(new[] { 12f, 12f, 136f }, Planner.PlanColumnWidths(Classifier.Classify(new List<ParagraphInfo>
        {
            new ParagraphInfo("1") { InTable = true, TableIndex = 1, CellRow = 1, CellColumn = 1 },
            new ParagraphInfo("2") { InTable = true, TableIndex = 1, CellRow = 1, CellColumn = 2 },
            new ParagraphInfo(new string('x', 80)) { InTable = true, TableIndex = 1, CellRow = 1, CellColumn = 3 },
        }, StandardKind.HanhChinh_ND30), 1, 160).Select(x => (float)System.Math.Round(x)).ToArray());
    }

    [Fact]
    public void ChapterNumberAndLowercaseChapterTitle()
    {
        var p = Profiles.PhapDien();
        Assert.Equal("Chương I", TextNormalizer.Fix(ElementKind.Chuong, "Chương i", p));
        Assert.Equal("Phần IV", TextNormalizer.Fix(ElementKind.Phan, "Phần iv", p));
        var cd = Classifier.Classify(Doc("Chương i", "quy định chung", "Điều 1. Phạm vi điều chỉnh"), p.Kind);
        Assert.Equal(new[] { ElementKind.Chuong, ElementKind.TenPhanChuong, ElementKind.Dieu }, cd.Kinds.ToArray());
        Assert.True(Planner.PlanParagraph(ElementKind.TenPhanChuong, cd.Paragraphs[1], p).ToUpper);
        Assert.True(Planner.NeedsTextFix(ElementKind.Chuong, p));
    }

    [Theory]
    [InlineData(ElementKind.ThongTinDau, "Tổ : Khoa học tự nhiên", "Tổ: Khoa học tự nhiên")]
    [InlineData(ElementKind.TrichYeu, "tên bài dạy : tốc độ chuyển động", "TÊN BÀI DẠY: tốc độ chuyển động")]
    [InlineData(ElementKind.TrichYeu, "Môn học : Khoa học tự nhiên ; lớp : 7", "Môn học: Khoa học tự nhiên; lớp: 7")]
    public void LessonPlanHeaderText(ElementKind kind, string input, string expected)
    {
        var p = Profiles.GiaoDuc();
        Assert.True(Planner.NeedsTextFix(kind, p));
        Assert.Equal(expected, TextNormalizer.Fix(kind, input, p));
        Assert.False(Planner.NeedsTextFix(ElementKind.TrichYeu, Profiles.HanhChinh()));
    }

    [Theory]
    [InlineData(ElementKind.NoiNhanItem, "- Lưu : VT, VP", true, "- Lưu: VT, VP.")]
    [InlineData(ElementKind.Body, "Căn cứ kế hoạch số 15 /KH-VP ngày 01 tháng 10", false, "Căn cứ kế hoạch số 15/KH-VP ngày 01 tháng 10")]
    [InlineData(ElementKind.Body, "Tỷ lệ 1 / 2 số hồ sơ", false, "Tỷ lệ 1/2 số hồ sơ")]
    public void ExtraSpacesAreRemoved(ElementKind kind, string input, bool isLast, string expected) =>
        Assert.Equal(expected, TextNormalizer.Fix(kind, input, Profiles.HanhChinh(), isLast));

    [Fact]
    public void LegalBasisOnlyBeforeFirstArticle()
    {
        var k = Classifier.Classify(Doc(QuyetDinh.Take(13).Concat(new[] { "Căn cứ kế hoạch số 15/KH-VP, các phòng tổ chức thực hiện." })
            .Concat(QuyetDinh.Skip(13)).ToArray()), StandardKind.HanhChinh_ND30).Kinds;
        Assert.Equal(ElementKind.CanCu, k[10]);    // căn cứ ban hành trước "QUYẾT ĐỊNH:"
        Assert.Equal(ElementKind.Dieu, k[12]);
        Assert.Equal(ElementKind.Body, k[13]);     // "Căn cứ..." sau Điều 1 là lời văn

        var dang = Classifier.Classify(Doc(BaoCaoDang.Take(10).Concat(new[] { "- Căn cứ kế hoạch của Huyện ủy, các chi bộ triển khai." }).ToArray()), StandardKind.Dang_HD05).Kinds;
        Assert.Equal(ElementKind.CanCu, dang[9]);
        var dang2 = Classifier.Classify(Doc(BaoCaoDang.Take(8).Concat(new[] { "I. TÌNH HÌNH CHUNG", "- Căn cứ kế hoạch của Huyện ủy, các chi bộ triển khai." }).ToArray()), StandardKind.Dang_HD05).Kinds;
        Assert.Equal(ElementKind.Body, dang2[9]);
    }

    [Fact]
    public void MessagesUseVietnameseDecimalSeparatorOnAnyCulture()
    {
        var saved = System.Globalization.CultureInfo.CurrentCulture;
        System.Globalization.CultureInfo.CurrentCulture = new System.Globalization.CultureInfo("en-US");
        try
        {
            var p = Profiles.HanhChinh();
            var issues = FormatChecker.Check(Classifier.Classify(Doc(QuyetDinh), p.Kind), p,
                new PageInfo { WidthMm = 210, HeightMm = 297, TopMm = 20, BottomMm = 20, LeftMm = 30, RightMm = 31.8f });
            Assert.Contains(issues, i => i.Message.Contains("31,8 mm"));
            Assert.Equal("en-US", System.Globalization.CultureInfo.CurrentCulture.Name);
        }
        finally { System.Globalization.CultureInfo.CurrentCulture = saved; }
    }
}
