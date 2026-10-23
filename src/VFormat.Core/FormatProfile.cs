using System.Collections.Generic;

namespace VFormat.Core
{
    /// <summary>Bộ tham số thể thức của một hệ quy chuẩn.</summary>
    public class FormatProfile
    {
        public StandardKind Kind;
        public string ShortName;      // "NĐ 30"
        public string FullName;       // "Nghị định 30/2020/NĐ-CP"
        public string Description;
        public string IssuePrefix;    // tiền tố mã lỗi: ND30, QD68, D05, DTN, PD250, GD5512

        // Trang A4 (mm)
        public float PageWidthMm = 210, PageHeightMm = 297;
        public float MarginTopMm = 20, MarginBottomMm = 20, MarginLeftMm = 30, MarginRightMm = 15;
        public float MinTopMm = 20, MaxTopMm = 25, MinBottomMm = 20, MaxBottomMm = 25;
        public float MinLeftMm = 30, MaxLeftMm = 35, MinRightMm = 15, MaxRightMm = 20;

        public string FontName = "Times New Roman";
        public float BodySize = 14;
        public float LineSpacing = 1.15f;     // bội số dòng (1.0 - 1.5)
        public float BodyFirstIndentCm = 1.0f;
        public float BodySpaceAfterPt = 6;

        // Số trang
        public bool PageNumberInHeader = true;
        public bool HidePageNumberOnFirstPage = true;
        public float PageNumberSize = 14;

        // Đặc thù thể thức
        public bool SoKyHieuHasColon = true;     // "Số: 12/QĐ-UBND" (Đảng: "Số 12-NQ/TU")
        public string SoPrefix = "Số:";
        public bool UseDauSao;                   // Đảng "*", Đoàn "***"
        public string DauSaoText = "*";
        public string HeaderTitle;               // "ĐẢNG CỘNG SẢN VIỆT NAM" thay cho Quốc hiệu
        public float CoQuanLineRatio = 0.4f;     // đường kẻ dưới tên cơ quan (0 = không kẻ)
        public float TrichYeuLineRatio = 0.4f;   // đường kẻ dưới trích yếu (0 = không kẻ)
        public float HeaderTitleLineRatio = 1f;  // đường kẻ dưới tiêu ngữ / tiêu đề Đảng-Đoàn
        public bool NoiNhanUnderline;            // Đảng/Đoàn: "Nơi nhận:" gạch chân
        public string ThayMatText = "TM.";       // Đảng: "T/M"
        public string NoiNhanSeparator = ";";    // Đảng: ","

        // Bố cục bảng khối đầu & chữ ký (mm)
        public float HeaderLeftColMm = 72, HeaderRightColMm = 93;   // Quốc hiệu cỡ 12 đậm ≈ 89 mm
        public float SignLeftColMm = 75, SignRightColMm = 90;

        public Dictionary<ElementKind, ElementStyle> Styles = new Dictionary<ElementKind, ElementStyle>();

        public ElementStyle StyleFor(ElementKind k)
        {
            if (Styles.TryGetValue(k, out var s)) return s;
            return Styles[ElementKind.Body];
        }

        public float TextWidthMm => PageWidthMm - MarginLeftMm - MarginRightMm;

        public FormatProfile Clone()
        {
            var p = (FormatProfile)MemberwiseClone();
            p.Styles = new Dictionary<ElementKind, ElementStyle>();
            foreach (var kv in Styles) p.Styles[kv.Key] = kv.Value.Clone();
            return p;
        }

        /// <summary>Đổi cỡ chữ nội dung (13 hoặc 14) và đồng bộ các thành phần theo nội dung.</summary>
        public void SetBodySize(float size)
        {
            BodySize = size;
            foreach (var k in new[] { ElementKind.Body, ElementKind.CanCu, ElementKind.Dieu, ElementKind.MucLaMa,
                                      ElementKind.MucSo, ElementKind.MucChuCai, ElementKind.KinhGui })
                if (Styles.TryGetValue(k, out var s)) s.Size = size;
        }
    }
}
