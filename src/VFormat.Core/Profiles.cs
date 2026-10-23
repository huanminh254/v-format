using System;
using System.Collections.Generic;
using System.Linq;

namespace VFormat.Core
{
    /// <summary>
    /// Bộ thông số mặc định cho 6 hệ quy chuẩn. Các giá trị lấy theo phụ lục thể thức của từng văn bản
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
            st[ElementKind.DiaDanhNgayThang] = S(14, 13, 14, italic: true, before: 6);
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

        public static FormatProfile QuanDoi()
        {
            var p = new FormatProfile
            {
                Kind = StandardKind.QuanDoi_TT68,
                ShortName = "TT 68",
                FullName = "Thông tư 68/2023/TT-BQP",
                Description = "Văn bản hành chính trong Quân đội",
                IssuePrefix = "QD68",
            };
            AddCommonStyles(p);
            // Quân đội: tên cơ quan cấp trên & cơ quan ban hành cỡ 13; người ký ghi cấp bậc trước họ tên.
            p.Styles[ElementKind.CoQuanChuQuan].Size = 13;
            p.Styles[ElementKind.CoQuanBanHanh].Size = 13;
            p.Styles[ElementKind.QuocHieu].Size = 13;
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
                MarginLeftMm = 30, MarginRightMm = 15,
                SoKyHieuHasColon = false,
                SoPrefix = "Số",
                UseDauSao = true,
                DauSaoText = "*",
                HeaderTitle = "ĐẢNG CỘNG SẢN VIỆT NAM",
                CoQuanLineRatio = 0,
                TrichYeuLineRatio = 0,
                NoiNhanUnderline = true,
                ThayMatText = "T/M",
                NoiNhanSeparator = ",",
                HeaderLeftColMm = 80, HeaderRightColMm = 85,
                BodySize = 14,
            };
            AddCommonStyles(p);
            var st = p.Styles;
            st[ElementKind.TieuDeDang] = S(15, 15, 15, bold: true, upper: true);
            st[ElementKind.CoQuanChuQuan] = S(14, 14, 14, upper: true);
            st[ElementKind.CoQuanBanHanh] = S(14, 14, 14, bold: true, upper: true);
            st[ElementKind.SoKyHieu] = S(14, 14, 14);
            st[ElementKind.DiaDanhNgayThang] = S(14, 14, 14, italic: true, align: Align.Right);
            st[ElementKind.TenLoai] = S(16, 15, 16, bold: true, upper: true, before: 18);
            st[ElementKind.TrichYeu] = S(14, 14, 15, bold: true);
            st[ElementKind.Body] = S(14, 14, 15, align: Align.Justify, indentCm: 1.0f, after: 6, single: false);
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
                TrichYeuLineRatio = 0.4f,
                NoiNhanUnderline = true,
                HeaderLeftColMm = 80, HeaderRightColMm = 85,
            };
            AddCommonStyles(p);
            var st = p.Styles;
            st[ElementKind.TieuDeDoan] = S(15, 15, 15, bold: true, upper: true, underline: false);
            st[ElementKind.CoQuanChuQuan] = S(14, 14, 14, upper: true);
            st[ElementKind.CoQuanBanHanh] = S(14, 14, 14, bold: true, upper: true);
            st[ElementKind.SoKyHieu] = S(14, 13, 14);
            st[ElementKind.DiaDanhNgayThang] = S(14, 14, 14, italic: true, align: Align.Right);
            st[ElementKind.TenLoai] = S(16, 15, 16, bold: true, upper: true, before: 18);
            st[ElementKind.NoiNhanTitle] = S(14, 14, 14, underline: true, align: Align.Left, before: 12);
            st[ElementKind.NoiNhanItem] = S(12, 12, 12, align: Align.Left);
            return p;
        }

        public static FormatProfile PhapDien()
        {
            var p = new FormatProfile
            {
                Kind = StandardKind.PhapDien_ND250,
                ShortName = "NĐ 250",
                FullName = "Pháp điển & hợp nhất văn bản QPPL",
                Description = "Văn bản hợp nhất, pháp điển quy phạm pháp luật",
                IssuePrefix = "PD250",
                LineSpacing = 1.0f,
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
            return p;
        }
    }
}
