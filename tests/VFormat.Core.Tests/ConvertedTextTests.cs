using System.Linq;
using VFormat.Core;
using Xunit;

public class ConvertedTextTests
{
    [Theory]
    [InlineData("BÁO CÁO ", "Kết quả sinh hoạt(3) Kết quả sinh hoạt CTVBM", LineBreakAction.Split)]
    [InlineData("CÁN BỘ BÁO CÁO ", "(Ký, ghi rõ họ tên)", LineBreakAction.Split)]
    [InlineData("IV. NHẬN XÉT, ĐỀ XUẤT: ", "1 Cấp trên trực tiếp ", LineBreakAction.Split)]
    [InlineData("1 Cấp trên trực tiếp ", "2 Đơn vị báo cáo ", LineBreakAction.Split)]
    [InlineData("Mẫu BL2 ", "BH theo TT số 44/2025/TT-BCA", LineBreakAction.Split)]
    [InlineData("Chức vụ: Cán bộ; Đơn vị: C08, Phòng Hướng dẫn tuần tra, kiểm soát giao thông ", "đường bộ, đường sắt", LineBreakAction.Join)]
    [InlineData("- Tình hình trật tự an toàn giao thông trên tuyến cao tốc Nghi Sơn - Diễn Châu - Bãi ", "Vọt – Bùng – Vạn Ninh", LineBreakAction.Join)]
    [InlineData("- 02 xe ô tô tải biển số 60H-302.35 và 18H-004.92 có dấu hiệu vi phạm dừng đỗ ", "không đúng quy định", LineBreakAction.Join)]
    [InlineData("CTVBM, chủ HTBM báo cáo, thống kê tài liệu, hiện vật do LLBM, CTVBM cung cấp). ", "Tháng 09/2026 cơ sở bí mật cung cấp 02 tin:", LineBreakAction.Split)]
    [InlineData("Tiếp tục nắm tình hình vi phạm trên tuyến cao tốc Nghi Sơn - Diễn Châu - Bãi Vọt – ", "- Bùng – Vạn Ninh", LineBreakAction.Split)]
    [InlineData("đã nêu ở trên. ", "", LineBreakAction.Remove)]
    public void DecidesLineBreaks(string before, string after, LineBreakAction expected) =>
        Assert.Equal(expected, ConvertedText.Decide(before, after));

    [Fact]
    public void ReflowJoinsWrappedLinesAndSplitsTitle()
    {
        var p = ConvertedText.Reflow("BÁO CÁO \vKết quả sinh hoạt CTVBM");
        Assert.Equal(new[] { "BÁO CÁO ", "Kết quả sinh hoạt CTVBM" }, p);
        p = ConvertedText.Reflow("- Việc tổ chức sinh hoạt tại HTBM “ TẠ”: Đảm bảo tuyệt đối an toàn, bí mật, không để lộ lọt \vthông tin.");
        Assert.Single(p);
        Assert.EndsWith("lộ lọt thông tin.", p[0]);
    }

    [Fact]
    public void DoubleLineBreakIsParagraphBoundary()
    {
        var p = ConvertedText.Reflow("Tập trung phát hiện các dấu hiệu vi phạm về nồng độ cồn, đi ngược chiều, \v……\v\v\tIII. KINH PHÍ ĐÃ CHI:");
        Assert.Equal(new[] { "Tập trung phát hiện các dấu hiệu vi phạm về nồng độ cồn, đi ngược chiều, ……", "III. KINH PHÍ ĐÃ CHI:" }, p);
        Assert.Equal(new[] { LineBreakAction.Join, LineBreakAction.Split, LineBreakAction.Remove },
            ConvertedText.Plan("Tập trung phát hiện các dấu hiệu vi phạm về nồng độ cồn, đi ngược chiều, \v……\v\v\tIII. KINH PHÍ ĐÃ CHI:"));
    }

    [Fact]
    public void SplitsMergedHeaderParagraph()
    {
        var raw = "1CỤC CẢNH SÁT GIAO      CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM \t\tTHÔNG\t Độc lập - Tự do - Hạnh phúc \v\t2PHÒNG 6\r";
        var lines = ConvertedText.SplitMergedHeader(raw);
        Assert.Equal(new[]
        {
            "CỤC CẢNH SÁT GIAO THÔNG\tCỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM",
            "PHÒNG 6\tĐộc lập - Tự do - Hạnh phúc",
        }, lines);
    }

    [Fact]
    public void MoneyAmountIsNotNumberedHeading()
    {
        var d = Classifier.Classify(new[] { "III. KINH PHÍ ĐÃ CHI:", "1.000.000đ", "1. Nhiệm vụ trọng tâm", "2.Kết quả" }
            .Select((t, i) => new ParagraphInfo(t) { Index = i }).ToList(), StandardKind.HanhChinh_ND30);
        Assert.Equal(ElementKind.Body, d.Kinds[1]);
        Assert.Equal(ElementKind.MucSo, d.Kinds[2]);
        Assert.Equal(ElementKind.MucSo, d.Kinds[3]);
    }

    [Fact]
    public void SignBlockAfterDigitalSignatureIsRecognised()
    {
        var lines = new[]
        {
            "CỤC CẢNH SÁT GIAO THÔNG\tCỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM", "PHÒNG 6\tĐộc lập - Tự do - Hạnh phúc",
            "BÁO CÁO", "Kết quả sinh hoạt", "Tôi: Nguyễn Quang Linh",
            "- Việc tổ chức sinh hoạt đảm bảo an toàn.", "V. Ý KIẾN CỦA LÃNH ĐẠO/CHỈ HUY CÓ THẨM QUYỀN: Thực hiện theo quy định.",
            "Hà Nội, ngày 21 tháng 09 năm 2026", "LÃNH ĐẠO/CHỈ HUY CÓ THẨM QUYỀN", "(Ký, ghi rõ họ tên và đóng dấu) VB21",
            "Ký bởi: Phạm Đức Duẩn", "11:46:03 +07:00", "Thượng tá Phạm Đức Duẩn",
        };
        // nội dung đủ dài như văn bản thật: ngày tháng của khối ký nằm ngoài vùng khối đầu
        lines = lines.Take(5).Concat(Enumerable.Repeat("Nội dung báo cáo chi tiết theo từng nhiệm vụ được giao.", 25)).Concat(lines.Skip(5)).ToArray();
        var paras = lines.Select((t, i) => new ParagraphInfo(t) { Index = i }).ToList();
        paras[35].InTable = true; paras[35].TableIndex = 1; paras[35].CellColumn = 2;
        var d = Classifier.Classify(paras, StandardKind.CongAn_BCA08);
        Assert.Equal(ElementKind.DiaDanhNgayThang, d.Kinds[32]);
        Assert.Equal(ElementKind.ChucVuKy, d.Kinds[33]);
        Assert.Equal(ElementKind.GhiChuKy, d.Kinds[34]);
        Assert.Equal(ElementKind.HoTenKy, d.Kinds[37]);
        // ngoài bảng: đặt vào nửa phải trang, không giữ thụt lề phải của bản gốc
        var plan = Planner.PlanParagraph(d, 33, Profiles.CongAn());
        Assert.Equal(Align.Center, plan.Align);
        // ngày tháng trong khối ký: căn lề theo quy chuẩn (Đảng: căn phải), vẫn ở nửa phải trang
        var hd05 = Classifier.Classify(paras, StandardKind.Dang_HD05);
        var datePlan = Planner.PlanParagraph(hd05, 32, Profiles.Get(StandardKind.Dang_HD05));
        Assert.Equal(Profiles.Get(StandardKind.Dang_HD05).StyleFor(ElementKind.DiaDanhNgayThang).Align, datePlan.Align);
        Assert.NotNull(datePlan.LeftIndentPt);
        Assert.Equal(Profiles.CongAn().SignLeftColMm * Planner.PtPerMm, plan.LeftIndentPt.Value, 1);
        Assert.NotNull(Planner.PlanParagraph(d, 32, Profiles.CongAn()).LeftIndentPt);
    }

    [Theory]
    [InlineData("CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM")]                       // Quốc hiệu đứng riêng
    [InlineData("Căn cứ Hiến pháp nước Cộng hòa xã hội chủ nghĩa Việt Nam;")] // lời văn
    public void LeavesNormalParagraphsAlone(string raw) => Assert.Null(ConvertedText.SplitMergedHeader(raw));

    [Fact]
    public void MergedHeaderIsClassifiedAsTwoColumnHeader()
    {
        var lines = ConvertedText.SplitMergedHeader(
            "1CỤC CẢNH SÁT GIAO      CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM \t\tTHÔNG\t Độc lập - Tự do - Hạnh phúc \v\t2PHÒNG 6");
        var paras = lines.Concat(new[] { "Hà Nội, ngày 20 tháng 9 năm 2026", "BÁO CÁO", "Kết quả sinh hoạt CTVBM" })
            .Select((t, i) => new ParagraphInfo(t) { Index = i }).ToList();
        var d = Classifier.Classify(paras, StandardKind.CongAn_BCA08);
        Assert.Equal(ElementKind.QuocHieu, d.Kinds[0]);
        Assert.Equal(ElementKind.TieuNgu, d.Kinds[1]);
        Assert.Equal(ElementKind.CoQuanChuQuan, d.Splits[0].LeftKind);
        Assert.Equal(ElementKind.CoQuanBanHanh, d.Splits[1].LeftKind);
        Assert.Equal(ElementKind.TenLoai, d.Kinds[3]);
    }
}
