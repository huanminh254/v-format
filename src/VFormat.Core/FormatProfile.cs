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
        public float LineSpacing = 1.15f;     // bội số dòng (1.0 - MaxLineSpacing)
        public float MaxLineSpacing = 1.5f;   // Đoàn: tối đa 1,2
        public float LineExactPt;             // > 0: giãn dòng Exactly (Đảng 18-22 pt) thay cho bội số
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
        public string KyThayText = "KT.";        // Đảng: "K/T"
        public string ThuaLenhText = "TL.";      // Đảng: "T/L"
        public string ThuaUyQuyenText = "TUQ.";  // Đảng: "T/UQ"
        public bool SignerTitleBold = true;      // Đảng, Đoàn: chức vụ người ký ("PHÓ BÍ THƯ") đứng, không đậm
        public bool KinhGuiPrefixItalic;         // Đảng, Đoàn: chữ "Kính gửi:" in nghiêng
        public string TitleDashes;               // Đảng "-----", Đoàn "-------" dưới tên loại, trích yếu
        public bool CanCuDash;                   // Đảng: "- Căn cứ ...;", căn cứ cuối kết thúc bằng dấu chấm
        public string NoiNhanSeparator = ";";
        public bool RankBeforeSignerName;        // Quân đội, Công an: cấp bậc hàm trước họ tên người ký
        public bool FormatAdvisory;              // CV 5512: chỉ quy định khung nội dung — lỗi lề, phông, cỡ chữ chỉ là khuyến nghị

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

        /// <summary>Các thành phần có cỡ chữ bằng cỡ chữ phần nội dung.</summary>
        public static readonly ElementKind[] BodySizedKinds =
        {
            ElementKind.Body, ElementKind.CanCu, ElementKind.Dieu, ElementKind.MucLaMa,
            ElementKind.MucSo, ElementKind.MucChuCai, ElementKind.KinhGui,
        };

        /// <summary>
        /// Đổi cỡ chữ nội dung (13 hoặc 14) và đồng bộ các thành phần theo nội dung. Cỡ ngoài khoảng quy chuẩn cho phép
        /// được đưa về giới hạn gần nhất (Đảng: nội dung 14–15, chọn 13 vẫn là 14).
        /// </summary>
        public void SetBodySize(float size)
        {
            foreach (var k in BodySizedKinds)
                if (Styles.TryGetValue(k, out var s)) s.Size = System.Math.Min(System.Math.Max(size, s.MinSize), s.MaxSize);
            BodySize = Styles.TryGetValue(ElementKind.Body, out var body) ? body.Size : size;
        }

        /// <summary>Đổi giãn dòng (bội số), không vượt mức tối đa của quy chuẩn.</summary>
        public void SetLineSpacing(float multiple) => LineSpacing = System.Math.Min(System.Math.Max(multiple, 1f), MaxLineSpacing);

        /// <summary>Đặt lề cố định (quy chuẩn không cho khoảng, ví dụ Đảng 20/20/30/15).</summary>
        public void SetFixedMargins(float top, float bottom, float left, float right)
        {
            MarginTopMm = MinTopMm = MaxTopMm = top;
            MarginBottomMm = MinBottomMm = MaxBottomMm = bottom;
            MarginLeftMm = MinLeftMm = MaxLeftMm = left;
            MarginRightMm = MinRightMm = MaxRightMm = right;
        }
    }
}
