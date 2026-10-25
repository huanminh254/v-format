using System.Collections.Generic;
using System.Linq;
using VFormat.Core;
using Xunit;

/// <summary>Các điểm sửa theo báo cáo đối chiếu 7 hệ quy chuẩn (docs/kiem-thu/2026-10-24-doi-chieu-7-quy-chuan.md).</summary>
public class QuyChuanTests
{
    static List<ParagraphInfo> Doc(params string[] lines) =>
        lines.Select((t, i) => new ParagraphInfo(t) { Index = i }).ToList();

    static ClassifiedDocument Classify(IList<ParagraphInfo> paras, StandardKind kind) => Classifier.Classify(paras, kind);

    // ---------- A1. Khối đầu gõ hai cột bằng tab ----------

    static readonly string[] BaoCaoT9 =
    {
        "CỤC CẢNH SÁT GIAO\tCỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM", "THÔNG\tĐộc lập - Tự do - Hạnh phúc",
        "Số: 12/BC-C08\tHà Nội, ngày 05 tháng 9 năm 2026", "",
        "BÁO CÁO", "Kết quả công tác tháng 9", "Nội dung báo cáo tổng hợp kết quả công tác.",
        "1.000.000đ là kinh phí đã chi cho hoạt động này trong tháng.",
        "Nơi nhận:", "- Lãnh đạo Cục;", "- Lưu: VT.", "CỤC TRƯỞNG", "", "", "Thiếu tướng Nguyễn Văn A",
    };

    [Fact]
    public void TabSeparatedHeaderIsRecognisedAndSplitIntoTwoColumns()
    {
        var p = Profiles.CongAn();
        var cd = Classify(Doc(BaoCaoT9), p.Kind);
        Assert.Equal(ElementKind.QuocHieu, cd.Kinds[0]);
        Assert.Equal(ElementKind.TieuNgu, cd.Kinds[1]);
        Assert.Equal(ElementKind.DiaDanhNgayThang, cd.Kinds[2]);
        // tên cơ quan bị xuống dòng ("... GIAO" / "THÔNG") được nối lại
        Assert.Equal("CỤC CẢNH SÁT GIAO THÔNG", cd.Splits[0].Left);
        Assert.Equal(ElementKind.CoQuanBanHanh, cd.Splits[0].LeftKind);
        Assert.Equal(ElementKind.Empty, cd.Splits[1].LeftKind);
        Assert.Equal(ElementKind.SoKyHieu, cd.Splits[2].LeftKind);

        var issues = FormatChecker.Check(cd, p);
        Assert.DoesNotContain(issues, x => x.Message.Contains("Thiếu Quốc hiệu"));
        Assert.DoesNotContain(issues, x => x.Message.Contains("Thiếu số, ký hiệu"));
        Assert.Contains(issues, x => x.Code == "CA08-07" && x.Message.Contains("tab"));

        var plan = Planner.PlanHeader(cd, p);
        Assert.NotNull(plan);
        Assert.Equal(new[] { "CỤC CẢNH SÁT GIAO THÔNG", "Số: 12/BC-C08" }, plan.Left);
        Assert.Equal(new[] { "CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM", "Độc lập - Tự do - Hạnh phúc", "Hà Nội, ngày 05 tháng 9 năm 2026" }, plan.Right);

        // chưa dàn bảng: dòng gõ tab căn trái, không in đậm cả dòng
        var pp = Planner.PlanParagraph(cd, 0, p);
        Assert.Equal(Align.Left, pp.Align);
        Assert.Null(pp.Bold);
        Assert.False(pp.ToUpper);
    }

    [Fact]
    public void DetectorReadsTabSeparatedHeader()
    {
        Assert.Equal(StandardKind.Dang_HD05, StandardDetector.Detect(Doc("HUYỆN ỦY ĐÔNG ANH\tĐẢNG CỘNG SẢN VIỆT NAM", "*", "Số 12-NQ/HU")));
    }

    // ---------- Trích yếu / A4 ----------

    [Fact]
    public void SentenceAfterSubjectIsBodyAndThousandsAreNotHeadings()
    {
        var cd = Classify(Doc(BaoCaoT9), StandardKind.CongAn_BCA08);
        Assert.Equal(ElementKind.TrichYeu, cd.Kinds[5]);
        Assert.Equal(ElementKind.Body, cd.Kinds[6]);
        Assert.Equal(ElementKind.Body, cd.Kinds[7]);
    }

    // ---------- A5. Ghi chú biểu mẫu ----------

    [Fact]
    public void FormNotesAreKeptAsIs()
    {
        var lines = new[] { "Mẫu BL2", "BH theo TT số 44/2025/TT-BCA", "Ngày 18/6/2025", "VB21" }.Concat(BaoCaoT9).ToArray();
        var p = Profiles.CongAn();
        var cd = Classify(Doc(lines), p.Kind);
        Assert.All(Enumerable.Range(0, 4), i => Assert.Equal(ElementKind.GhiChuMau, cd.Kinds[i]));
        Assert.Equal(ElementKind.QuocHieu, cd.Kinds[4]);
        var plan = Planner.PlanParagraph(cd, 3, p);
        Assert.Null(plan.Size);
        Assert.Null(plan.Bold);
        Assert.Null(plan.Align);
        Assert.DoesNotContain(FormatChecker.Check(cd, p), x => x.ParagraphIndex >= 0 && x.ParagraphIndex < 4);
    }

    [Fact]
    public void SmallPrintAtTopIsFormNote()
    {
        var paras = Doc(new[] { "Phiếu số 3 dùng cho đơn vị" }.Concat(BaoCaoT9).ToArray());
        paras[0].FontSize = 8;
        Assert.Equal(ElementKind.GhiChuMau, Classify(paras, StandardKind.CongAn_BCA08).Kinds[0]);
    }

    // ---------- A3, A6. Danh sách đánh số tự động ----------

    [Fact]
    public void AutoNumberedHeadingsUseWordListNumber()
    {
        var paras = Doc(BaoCaoT9);
        paras.Insert(7, new ParagraphInfo("KẾT QUẢ ĐẠT ĐƯỢC") { ListPrefix = "II.", InList = true });
        paras.Insert(8, new ParagraphInfo("THU NỘP NGÂN SÁCH") { ListPrefix = "•", InList = true });
        for (int i = 0; i < paras.Count; i++) paras[i].Index = i;
        var cd = Classify(paras, StandardKind.CongAn_BCA08);
        Assert.Equal(ElementKind.MucLaMa, cd.Kinds[7]);
        Assert.NotEqual(ElementKind.TenPhanChuong, cd.Kinds[8]);
    }

    [Fact]
    public void ListParagraphIndentIsNotReported()
    {
        var paras = Doc(BaoCaoT9);
        paras[6].InList = true; paras[6].ListPrefix = "-"; paras[6].FirstIndentPt = -0.6f;
        paras[6].Alignment = Align.Justify; paras[6].FontName = "Times New Roman"; paras[6].FontSize = 14;
        var issues = FormatChecker.Check(Classify(paras, StandardKind.CongAn_BCA08), Profiles.CongAn());
        Assert.DoesNotContain(issues, x => x.ParagraphIndex == 6 && x.Message.Contains("Thụt đầu dòng"));
    }

    // ---------- A2. Khối chữ ký đặt bằng thụt lề ----------

    [Fact]
    public void IndentedSignatureBlocksAreRecognisedAndKeepTheirIndent()
    {
        var paras = Doc(BaoCaoT9.Take(8).Concat(new[]
        {
            "CÁN BỘ BÁO CÁO", "(Ký, ghi rõ họ tên)", "", "Thiếu tá Nguyễn Văn B",
            "V. Ý KIẾN CỦA LÃNH ĐẠO/CHỈ HUY CÓ THẨM QUYỀN", "Đồng ý với nội dung báo cáo.",
            "LÃNH ĐẠO/CHỈ HUY CÓ THẨM QUYỀN", "(Ký, ghi rõ họ tên, đóng dấu)", "", "Đại tá Trần Văn C",
        }).ToArray());
        foreach (var i in new[] { 8, 9, 11, 14, 15, 17 }) paras[i].LeftIndentPt = 283;   // 10 cm
        var p = Profiles.CongAn();
        var cd = Classify(paras, p.Kind);
        Assert.Equal(ElementKind.ChucVuKy, cd.Kinds[8]);
        Assert.Equal(ElementKind.GhiChuKy, cd.Kinds[9]);
        Assert.Equal(ElementKind.HoTenKy, cd.Kinds[11]);
        Assert.Equal(ElementKind.MucLaMa, cd.Kinds[12]);
        Assert.Equal(ElementKind.ChucVuKy, cd.Kinds[14]);
        Assert.Equal(ElementKind.GhiChuKy, cd.Kinds[15]);
        Assert.Equal(ElementKind.HoTenKy, cd.Kinds[17]);

        foreach (var i in new[] { 8, 9, 11 })
        {
            var plan = Planner.PlanParagraph(cd, i, p);
            Assert.False(plan.ResetSideIndents);
            Assert.Equal(0, plan.FirstIndentPt);
            Assert.Equal(Align.Center, plan.Align);
        }
    }

    // ---------- Công an: cấp bậc hàm ----------

    [Theory]
    [InlineData("Thiếu tướng Nguyễn Văn A", true)]
    [InlineData("Đại úy Lê Văn B", true)]
    [InlineData("Thượng sĩ Trần Văn C", true)]
    [InlineData("Trung sĩ Trần Văn C", true)]
    [InlineData("Hạ sĩ Phạm Văn D", true)]
    [InlineData("Binh nhất Hoàng Văn E", true)]
    [InlineData("PGS.TS Thiếu tướng Nguyễn Văn A", true)]
    [InlineData("GS, TS Đại tá Nguyễn Văn A", true)]
    [InlineData("Nguyễn Văn A", false)]
    [InlineData("PGS.TS Nguyễn Văn A", false)]
    public void PoliceRanksIncludeNcosAndAcademicTitles(string signer, bool hasRank)
    {
        Assert.Equal(hasRank, Classifier.HasRank(signer));
        Assert.True(Classifier.IsPersonName(signer));
    }

    // ---------- NĐ 250/2026: văn bản hợp nhất ----------

    [Fact]
    public void Nd250ProfileNamesTheDecreeWithYear()
    {
        var p = Profiles.PhapDien();
        Assert.Equal("Nghị định 250/2026/NĐ-CP", p.FullName);
        Assert.Equal("NĐ 250/2026", p.ShortName);
        Assert.Equal(0, p.CoQuanLineRatio);
    }

    [Fact]
    public void ConsolidatedLabelMovesIntoHeaderLeftCell()
    {
        var p = Profiles.PhapDien();
        var cd = Classify(Doc("BỘ TÀI CHÍNH", "Số: 1/VBHN-BTC", "CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM",
            "Độc lập - Tự do - Hạnh phúc", "Hà Nội, ngày 01 tháng 3 năm 2026", "VĂN BẢN HỢP NHẤT", "THÔNG TƯ", "Quy định về kế toán",
            "Chương I", "QUY ĐỊNH CHUNG"), p.Kind);
        var plan = Planner.PlanHeader(cd, p);
        Assert.Equal(new[] { "BỘ TÀI CHÍNH", "VĂN BẢN HỢP NHẤT", "Số: 1/VBHN-BTC" }, plan.Left);
        Assert.Equal(new[] { 5 }, plan.Remove);

        var issues = FormatChecker.Check(cd, p);
        Assert.Contains(issues, x => x.ParagraphIndex == 5 && x.Message.Contains("ô trái khối đầu"));
        Assert.Contains(issues, x => x.Code == "PD250-08" && x.Message.Contains("năm ký xác thực"));
        Assert.Contains(issues, x => x.Message.Contains("XÁC THỰC VĂN BẢN HỢP NHẤT"));
        // số nhỏ hơn 10 thêm số 0
        Assert.Equal("Số: 01/2026/VBHN-BTC", TextNormalizer.FixSoKyHieu("Số: 1/2026/VBHN-BTC", p));
    }

    [Fact]
    public void ConsolidatedLabelInsideHeaderTableIsMovedAboveNumber()
    {
        var paras = Doc("BỘ TÀI CHÍNH", "Số: 01/2026/VBHN-BTC", "CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM", "Độc lập - Tự do - Hạnh phúc",
            "Hà Nội, ngày 01 tháng 3 năm 2026", "VĂN BẢN HỢP NHẤT", "THÔNG TƯ", "Quy định về kế toán");
        for (int i = 0; i < 5; i++) { paras[i].InTable = true; paras[i].TableIndex = 1; paras[i].CellRow = 1; paras[i].CellColumn = i < 2 ? 1 : 2; }
        var cd = Classify(paras, StandardKind.PhapDien_ND250);
        Assert.Equal((5, 1), Planner.PlanConsolidatedLabelMove(cd, Profiles.PhapDien()));
    }

    static readonly string[] HopNhat =
    {
        "VĂN BẢN HỢP NHẤT", "Số: 01/2026/VBHN-BTP", "CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM", "Độc lập - Tự do - Hạnh phúc",
        "Hà Nội, ngày 15 tháng 7 năm 2026", "",
        "THÔNG TƯ", "Quy định về hộ tịch",
        "Thông tư số 04/2020/TT-BTP ngày 28 tháng 5 năm 2020 của Bộ trưởng Bộ Tư pháp, có hiệu lực kể từ ngày 16 tháng 7 năm 2020, được sửa đổi, bổ sung bởi:",
        "Chương I", "QUY ĐỊNH CHUNG", "Điều 1. Phạm vi điều chỉnh", "Thông tư này quy định chi tiết một số điều của Luật Hộ tịch.",
        "Nơi nhận:", "- Văn phòng Chính phủ (để đăng Công báo);", "- Lưu: VT, HTQTCT.",
        "BỘ TƯ PHÁP", "XÁC THỰC VĂN BẢN HỢP NHẤT", "Hà Nội, ngày 15 tháng 7 năm 2026", "BỘ TRƯỞNG", "(Chữ ký và dấu)", "", "Nguyễn Văn A",
    };

    [Fact]
    public void ConsolidatedAuthenticationBlock()
    {
        var p = Profiles.PhapDien();
        var cd = Classify(Doc(HopNhat), p.Kind);
        Assert.Equal(ElementKind.CoQuanBanHanh, cd.Kinds[0]);
        Assert.Equal(ElementKind.TenLoai, cd.Kinds[6]);
        Assert.Equal(new[] { ElementKind.CoQuanXacThuc, ElementKind.XacThuc, ElementKind.DiaDanhNgayThang, ElementKind.ChucVuKy,
            ElementKind.GhiChuKy, ElementKind.Empty, ElementKind.HoTenKy }, cd.Kinds.Skip(16).ToArray());

        var sign = Planner.PlanSign(cd, p);
        Assert.NotNull(sign);
        Assert.Equal(new[] { "BỘ TƯ PHÁP", "XÁC THỰC VĂN BẢN HỢP NHẤT", "Hà Nội, ngày 15 tháng 7 năm 2026", "BỘ TRƯỞNG", "(Chữ ký và dấu)" },
            sign.Right.Take(5));
        Assert.Equal("Nguyễn Văn A", sign.Right.Last());

        // đường kẻ ngang nét liền dưới đoạn nội dung cuối, ngăn phần ký xác thực (mục XIII)
        Assert.Equal(12, Planner.SignSeparatorIndex(cd, p));
        Assert.True(Planner.PlanParagraph(cd, 12, p).BorderBottom);
        Assert.False(Planner.PlanParagraph(cd, 11, p).BorderBottom);
        Assert.Contains(Planner.PlanLines(cd, p), x => x.index == 16);
        Assert.DoesNotContain(Planner.PlanLines(cd, p), x => x.index == 0);   // không kẻ dưới "VĂN BẢN HỢP NHẤT"

        var issues = FormatChecker.Check(cd, p);
        Assert.DoesNotContain(issues, x => x.Message.Contains("VĂN BẢN HỢP NHẤT") || x.Message.Contains("VBHN") || x.Message.Contains("xác thực"));
    }

    [Fact]
    public void Nd250TemplateFollowsAppendix()
    {
        var t = TemplateLibrary.All.Single(x => x.Id == "pd-hn");
        var p = Profiles.PhapDien();
        var cd = Classify(Doc(t.Lines.ToArray()), p.Kind);
        Assert.Equal("VĂN BẢN HỢP NHẤT", t.Lines[0]);
        Assert.True(cd.Has(ElementKind.XacThuc));
        Assert.True(cd.Has(ElementKind.CoQuanXacThuc));
        Assert.DoesNotContain(FormatChecker.Check(cd, p), x => x.Severity == Severity.Error);
    }

    // ---------- CV 5512: khung kế hoạch bài dạy ----------

    [Fact]
    public void LessonPlanTemplatePassesFrameworkCheck()
    {
        var t = TemplateLibrary.All.Single(x => x.Id == "gd-khbd");
        Assert.Contains("I. Mục tiêu", t.Lines);
        Assert.Contains("II. Thiết bị dạy học và học liệu", t.Lines);
        Assert.Contains(t.Lines, x => x.StartsWith("2. Hoạt động 2: Hình thành kiến thức mới/giải quyết vấn đề/thực thi nhiệm vụ đặt ra từ Hoạt động 1"));
        var cd = Classify(Doc(t.Lines.ToArray()), StandardKind.GiaoDuc_CV5512);
        Assert.DoesNotContain(FormatChecker.Check(cd, Profiles.GiaoDuc()), x => x.Code == "GD5512-11");
    }

    [Fact]
    public void LessonPlanMissingPartsAreReported()
    {
        var t = TemplateLibrary.All.Single(x => x.Id == "gd-khbd");
        var lines = t.Lines.ToList();
        int act2 = lines.FindIndex(x => x.StartsWith("2. Hoạt động 2"));
        lines.RemoveAt(lines.FindIndex(act2, x => x == "c) Sản phẩm"));
        int act4 = lines.FindIndex(x => x.StartsWith("4. Hoạt động 4"));
        lines.RemoveRange(act4, lines.Count - act4);
        lines.Remove("II. Thiết bị dạy học và học liệu");

        var issues = FormatChecker.Check(Classify(Doc(lines.ToArray()), StandardKind.GiaoDuc_CV5512), Profiles.GiaoDuc())
            .Where(x => x.Code == "GD5512-11").ToList();
        Assert.Contains(issues, x => x.Message.Contains("Hoạt động 2 thiếu: c) Sản phẩm"));
        Assert.Contains(issues, x => x.Message.Contains("II. Thiết bị dạy học và học liệu") && x.Severity == Severity.Error);
        // hoạt động vận dụng chỉ làm với bài phù hợp (Ghi chú 1): chỉ nhắc
        Assert.Contains(issues, x => x.Message.Contains("Hoạt động 4") && x.Severity == Severity.Info);
    }

    [Fact]
    public void LessonPlanFormattingIssuesAreOnlyAdvice()
    {
        var t = TemplateLibrary.All.Single(x => x.Id == "gd-khbd");
        var paras = Doc(t.Lines.ToArray());
        paras[4].FontName = "Arial";
        var issues = FormatChecker.Check(Classify(paras, StandardKind.GiaoDuc_CV5512), Profiles.GiaoDuc(),
            new PageInfo { WidthMm = 210, HeightMm = 297, TopMm = 15, BottomMm = 15, LeftMm = 20, RightMm = 10 });
        var format = issues.Where(x => x.Code != "GD5512-11").ToList();
        Assert.NotEmpty(format);
        Assert.All(format, x => { Assert.Equal(Severity.Info, x.Severity); Assert.Contains("khuyến nghị", x.Message); });
    }
}
