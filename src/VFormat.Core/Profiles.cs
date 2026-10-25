using System;
using System.Collections.Generic;
using System.Linq;

namespace VFormat.Core
{
    /// <summary>
    /// Bộ thông số mặc định cho 7 hệ quy chuẩn. Các giá trị lấy theo phụ lục thể thức của từng văn bản
    /// quy định; khi quy định cho phép một khoảng (ví dụ cỡ 13-14) thì Size là giá trị mặc định
    /// còn MinSize/MaxSize là khoảng chấp nhận khi rà soát.
    /// </summary>
    public static class Profiles
    {
        public static readonly StandardKind[] All =
            (StandardKind[])Enum.GetValues(typeof(StandardKind));

        public static FormatProfile Get(StandardKind kind)
        {
            switch (kind)
            {
                case StandardKind.QuanDoi_TT68: return QuanDoi();
                case StandardKind.Dang_HD05: return Dang();
                case StandardKind.Doan_HD29: return Doan();
                case StandardKind.PhapDien_ND250: return PhapDien();
                case StandardKind.GiaoDuc_CV5512: return GiaoDuc();
                case StandardKind.CongAn_BCA08: return CongAn();
                default: return HanhChinh();
            }
        }

        public static IEnumerable<FormatProfile> GetAll() => All.Select(Get);

        static ElementStyle S(float size, float min, float max, bool bold = false, bool italic = false,
            bool upper = false, Align align = Align.Center, float indentCm = 0, float before = 0, float after = 0,
            bool single = true, bool underline = false)
            => new ElementStyle
            {
                Size = size, MinSize = min, MaxSize = max, Bold = bold, Italic = italic, UpperCase = upper,
                Align = align, FirstIndentCm = indentCm, SpaceBeforePt = before, SpaceAfterPt = after,
                KeepLineSpacingSingle = single, Underline = underline,
            };

        /// <summary>Thông số chung cho văn bản hành chính (Phụ lục I NĐ 30/2020).</summary>
        static void AddCommonStyles(FormatProfile p)
        {
            float b = p.BodySize;
            var st = p.Styles;
            st[ElementKind.Body] = S(b, 13, 14, align: Align.Justify, indentCm: p.BodyFirstIndentCm, after: p.BodySpaceAfterPt, single: false);
            st[ElementKind.CanCu] = S(b, 13, 14, italic: true, align: Align.Justify, indentCm: p.BodyFirstIndentCm, after: p.BodySpaceAfterPt, single: false);
            st[ElementKind.KinhGui] = S(b, 13, 14, align: Align.Center, before: 6, after: 6);
            st[ElementKind.QuocHieu] = S(12, 12, 13, bold: true, upper: true);
            st[ElementKind.TieuNgu] = S(13, 13, 14, bold: true);
            st[ElementKind.CoQuanChuQuan] = S(12, 12, 13, upper: true);
            st[ElementKind.CoQuanBanHanh] = S(12, 12, 13, bold: true, upper: true);
            st[ElementKind.DauSao] = S(12, 12, 14);
            st[ElementKind.SoKyHieu] = S(13, 13, 13, before: 6);
            st[ElementKind.TrichYeuCongVan] = S(12, 12, 13);
            // cỡ chữ khối đầu thống nhất (NĐ 30 Phụ lục I, chú thích mục V): Quốc hiệu 12 → Tiêu ngữ 13 → ngày tháng 13
            st[ElementKind.DiaDanhNgayThang] = S(13, 13, 14, italic: true, before: 6);
            st[ElementKind.TenLoai] = S(14, 13, 14, bold: true, upper: true, before: 18);
            st[ElementKind.TrichYeu] = S(14, 13, 14, bold: true, after: 0);
            st[ElementKind.Phan] = S(14, 13, 14, bold: true, before: 12);
            st[ElementKind.Chuong] = S(14, 13, 14, bold: true, before: 12);
            st[ElementKind.Muc] = S(14, 13, 14, bold: true, before: 6);
            st[ElementKind.TenPhanChuong] = S(14, 13, 14, bold: true, upper: true, after: 6);
            st[ElementKind.Dieu] = S(b, 13, 14, bold: true, align: Align.Justify, indentCm: p.BodyFirstIndentCm, before: 6, after: p.BodySpaceAfterPt, single: false);
            st[ElementKind.MucLaMa] = S(b, 13, 14, bold: true, align: Align.Justify, indentCm: p.BodyFirstIndentCm, before: 6, after: p.BodySpaceAfterPt, single: false);
            st[ElementKind.MucSo] = S(b, 13, 14, bold: true, align: Align.Justify, indentCm: p.BodyFirstIndentCm, after: p.BodySpaceAfterPt, single: false);
            st[ElementKind.MucChuCai] = S(b, 13, 14, align: Align.Justify, indentCm: p.BodyFirstIndentCm, after: p.BodySpaceAfterPt, single: false);
            st[ElementKind.TableText] = S(13, 11, 14, align: Align.Left);
            st[ElementKind.ChucVuKy] = S(14, 13, 14, bold: true, upper: true);
            st[ElementKind.HoTenKy] = S(14, 13, 14, bold: true);
            st[ElementKind.NoiNhanTitle] = S(12, 12, 12, bold: true, italic: true, align: Align.Left, before: 12);
            st[ElementKind.NoiNhanItem] = S(11, 11, 11, align: Align.Left);
            st[ElementKind.Empty] = S(b, 11, 14, align: Align.Justify);
            st[ElementKind.ThongTinDau] = S(b, 13, 14, align: Align.Left);
            st[ElementKind.GhiChuMau] = S(10, 6, 14, align: Align.Right);
            st[ElementKind.GhiChuKy] = S(b, 12, 14, italic: true);
            // phần ký xác thực văn bản hợp nhất (NĐ 250/2026 Phụ lục mục XIII)
            st[ElementKind.CoQuanXacThuc] = S(13, 12, 13, bold: true, upper: true);
            st[ElementKind.XacThuc] = S(14, 13, 14, bold: true, upper: true);
        }

        public static FormatProfile HanhChinh()
        {
            var p = new FormatProfile
            {
                Kind = StandardKind.HanhChinh_ND30,
                ShortName = "NĐ 30",
                FullName = "Nghị định 30/2020/NĐ-CP",
                Description = "Văn bản hành chính Nhà nước",
                IssuePrefix = "ND30",
            };
            AddCommonStyles(p);
            return p;
        }

        /// <summary>
        /// Văn bản hợp nhất 08/VBHN-BCA (hợp nhất Thông tư 26/2019/TT-BCA): Điều 9 dẫn chiếu thể thức và kỹ thuật
        /// trình bày theo Điều 8, 9 NĐ 30/2020; phần riêng của CAND là ghi cấp bậc hàm trước họ tên người ký.
        /// </summary>
        public static FormatProfile CongAn()
        {
            var p = new FormatProfile
            {
                Kind = StandardKind.CongAn_BCA08,
                ShortName = "VBHN 08",
                FullName = "Văn bản hợp nhất 08/VBHN-BCA",
                Description = "Văn bản Công an nhân dân",
                IssuePrefix = "CA08",
                RankBeforeSignerName = true,
            };
            AddCommonStyles(p);
            return p;
        }

        public static FormatProfile QuanDoi()
        {
            var p = new FormatProfile
            {
                Kind = StandardKind.QuanDoi_TT68,
                ShortName = "TT 68",
                FullName = "Thông tư 68/2023/TT-BQP",
                Description = "Văn bản hành chính trong Quân đội",
                IssuePrefix = "QD68",
                RankBeforeSignerName = true,
            };
            AddCommonStyles(p);
            // Quân đội: tên cơ quan cấp trên & cơ quan ban hành cỡ 13; người ký ghi cấp bậc trước họ tên.
            p.Styles[ElementKind.CoQuanChuQuan].Size = 13;
            p.Styles[ElementKind.CoQuanBanHanh].Size = 13;
            p.Styles[ElementKind.QuocHieu].Size = 13;
            p.Styles[ElementKind.TieuNgu].Size = 14;            // Quốc hiệu 13 → Tiêu ngữ 14 → ngày tháng 14
            p.Styles[ElementKind.DiaDanhNgayThang].Size = 14;
            p.HeaderRightColMm = 100;   // Quốc hiệu cỡ 13 đậm ≈ 96 mm
            p.HeaderLeftColMm = 68;
            return p;
        }

        public static FormatProfile Dang()
        {
            var p = new FormatProfile
            {
                Kind = StandardKind.Dang_HD05,
                ShortName = "HD 05",
                FullName = "Hướng dẫn 05-HD/VPTW",
                Description = "Văn bản của Đảng",
                IssuePrefix = "D05",
                SoKyHieuHasColon = false,
                SoPrefix = "Số",
                UseDauSao = true,
                DauSaoText = "*",
                HeaderTitle = "ĐẢNG CỘNG SẢN VIỆT NAM",
                CoQuanLineRatio = 0,
                TrichYeuLineRatio = 0,
                NoiNhanUnderline = true,
                ThayMatText = "T/M", KyThayText = "K/T", ThuaLenhText = "T/L", ThuaUyQuyenText = "T/UQ",
                SignerTitleBold = false,
                KinhGuiPrefixItalic = true,
                TitleDashes = "-----",
                CanCuDash = true,
                LineExactPt = 18,    // HD 05/2026 mục 6.2.2: Exactly 18-22 pt
                HeaderLeftColMm = 80, HeaderRightColMm = 85,
                BodySize = 14,
            };
            p.SetFixedMargins(20, 20, 30, 15);   // HD 05/2026 mục IV.4: lề cố định, không phải khoảng
            AddCommonStyles(p);
            var st = p.Styles;
            st[ElementKind.TieuDeDang] = S(15, 15, 15, bold: true, upper: true);
            st[ElementKind.CoQuanChuQuan] = S(14, 14, 14, upper: true);
            st[ElementKind.CoQuanBanHanh] = S(14, 14, 14, bold: true, upper: true);
            st[ElementKind.SoKyHieu] = S(14, 14, 14);
            st[ElementKind.TrichYeuCongVan] = S(12, 12, 12, italic: true);
            st[ElementKind.DiaDanhNgayThang] = S(14, 14, 14, italic: true, align: Align.Right);
            st[ElementKind.TenLoai] = S(16, 15, 16, bold: true, upper: true, before: 18);
            st[ElementKind.TrichYeu] = S(14, 14, 15, bold: true);
            st[ElementKind.Body] = S(14, 14, 15, align: Align.Justify, indentCm: 1.0f, after: 6, single: false);
            // căn cứ: kiểu chữ đứng như nội dung (HD 05 không quy định in nghiêng như NĐ 30)
            st[ElementKind.CanCu] = S(14, 14, 15, align: Align.Justify, indentCm: 1.0f, after: 6, single: false);
            // khoản "1. ...": đứng, không đậm (Phụ lục 2 mục 6)
            st[ElementKind.MucSo] = S(14, 14, 15, align: Align.Justify, indentCm: 1.0f, after: 6, single: false);
            // phần nội dung văn bản Đảng cỡ 14–15 (không phải 13–14 như NĐ 30), kể cả phần, chương, mục và tên
            foreach (var k in FormatProfile.BodySizedKinds.Concat(new[] { ElementKind.Phan, ElementKind.Chuong, ElementKind.Muc, ElementKind.TenPhanChuong }))
            { st[k].MinSize = 14; st[k].MaxSize = 15; }
            st[ElementKind.NoiNhanTitle] = S(14, 14, 14, underline: true, align: Align.Left, before: 12);
            st[ElementKind.NoiNhanItem] = S(12, 12, 12, align: Align.Left);
            st[ElementKind.ChucVuKy] = S(14, 14, 14, bold: true, upper: true);
            st[ElementKind.HoTenKy] = S(14, 14, 14, bold: true);
            return p;
        }

        public static FormatProfile Doan()
        {
            var p = new FormatProfile
            {
                Kind = StandardKind.Doan_HD29,
                ShortName = "HD 29",
                FullName = "Hướng dẫn 29-HD/TWĐTN-VP",
                Description = "Văn bản Đoàn TNCS Hồ Chí Minh",
                IssuePrefix = "DTN",
                SoKyHieuHasColon = true,
                UseDauSao = true,
                DauSaoText = "***",
                HeaderTitle = "ĐOÀN TNCS HỒ CHÍ MINH",
                CoQuanLineRatio = 0,
                TrichYeuLineRatio = 0,
                TitleDashes = "-------",     // gạch nét rời dưới tên loại, trích yếu (mục 5)
                SignerTitleBold = false,
                KinhGuiPrefixItalic = true,
                BodyFirstIndentCm = 1.27f,   // lùi đầu dòng 1 default tab (mục 6)
                MaxLineSpacing = 1.2f,       // giãn dòng từ dòng đơn tới tối đa 1,2
                HeaderLeftColMm = 80, HeaderRightColMm = 85,
            };
            p.SetFixedMargins(20, 20, 30, 20);   // mục II.3: lề phải 20 mm
            AddCommonStyles(p);
            var st = p.Styles;
            // Mẫu 3: kiểu chữ, cỡ chữ từng thành phần thể thức
            st[ElementKind.TieuDeDoan] = S(15, 15, 15, bold: true, upper: true, underline: false);
            st[ElementKind.CoQuanChuQuan] = S(14, 14, 14, upper: true);
            st[ElementKind.CoQuanBanHanh] = S(14, 14, 14, bold: true, upper: true);
            st[ElementKind.SoKyHieu] = S(14, 13, 14);
            st[ElementKind.DiaDanhNgayThang] = S(13, 13, 13, italic: true, align: Align.Right);
            st[ElementKind.TrichYeuCongVan] = S(12, 12, 12, italic: true);
            st[ElementKind.TenLoai] = S(16, 15, 16, bold: true, upper: true, before: 18);
            st[ElementKind.KinhGui] = S(14, 14, 14, bold: true, align: Align.Center, before: 6, after: 6);
            st[ElementKind.Phan] = S(15, 15, 16, bold: true, before: 12);
            st[ElementKind.TenPhanChuong] = S(14, 14, 14, bold: true, after: 6);   // tiêu đề phần: chữ thường, đậm
            st[ElementKind.NoiNhanTitle] = S(13, 13, 13, bold: true, align: Align.Left, before: 12);
            st[ElementKind.NoiNhanItem] = S(11, 11, 11, align: Align.Left);
            // phần lời văn cỡ 14 cố định (mục 6)
            foreach (var k in FormatProfile.BodySizedKinds) { st[k].MinSize = 14; st[k].MaxSize = 14; }
            return p;
        }

        public static FormatProfile PhapDien()
        {
            var p = new FormatProfile
            {
                Kind = StandardKind.PhapDien_ND250,
                ShortName = "NĐ 250/2026",
                FullName = "Nghị định 250/2026/NĐ-CP",
                Description = "Kỹ thuật trình bày văn bản hợp nhất",
                IssuePrefix = "PD250",
                LineSpacing = 1.0f,
                // ô trái khối đầu là "VĂN BẢN HỢP NHẤT" + số, ký hiệu (Phụ lục mục I.2, ví dụ 2): không kẻ dưới
                CoQuanLineRatio = 0,
            };
            AddCommonStyles(p);
            var st = p.Styles;
            st[ElementKind.Phan] = S(14, 14, 14, bold: true, before: 12);
            st[ElementKind.Chuong] = S(14, 14, 14, bold: true, before: 12);
            st[ElementKind.Muc] = S(14, 14, 14, bold: true, before: 6);
            st[ElementKind.TenPhanChuong] = S(14, 14, 14, bold: true, upper: true, after: 6);
            st[ElementKind.MucSo] = S(14, 13, 14, align: Align.Justify, indentCm: 1.0f, after: 6, single: false); // khoản: không đậm
            return p;
        }

        public static FormatProfile GiaoDuc()
        {
            var p = new FormatProfile
            {
                Kind = StandardKind.GiaoDuc_CV5512,
                ShortName = "CV 5512",
                FullName = "Công văn 5512/BGDĐT-GDTrH",
                Description = "Kế hoạch bài dạy ngành Giáo dục",
                IssuePrefix = "GD5512",
                MarginTopMm = 20, MarginBottomMm = 20, MarginLeftMm = 30, MarginRightMm = 20,
                MaxRightMm = 20, BodyFirstIndentCm = 0.75f, LineSpacing = 1.15f,
                TrichYeuLineRatio = 0, // Phụ lục IV không kẻ dưới tên bài
            };
            AddCommonStyles(p);
            var st = p.Styles;
            st[ElementKind.TenLoai] = S(14, 14, 16, bold: true, upper: true, before: 6);
            st[ElementKind.TrichYeu] = S(14, 13, 14); // "Môn học...; lớp..." đứng thường; dòng in hoa "TÊN BÀI DẠY" in đậm
            st[ElementKind.MucLaMa] = S(14, 13, 14, bold: true, align: Align.Left, before: 6, after: 3, single: false);
            st[ElementKind.MucSo] = S(14, 13, 14, bold: true, align: Align.Left, indentCm: 0, after: 3, single: false);
            st[ElementKind.MucChuCai] = S(14, 13, 14, bold: true, italic: true, align: Align.Justify, indentCm: 0.75f, after: 3, single: false);
            p.PageNumberInHeader = false; // KHBD thường đánh số ở chân trang
            // Phụ lục IV chỉ quy định khung nội dung; thông số trình bày trên là khuyến nghị của V-FORMAT
            p.FormatAdvisory = true;
            return p;
        }
    }
}
