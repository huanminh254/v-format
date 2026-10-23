using System.Collections.Generic;

namespace VFormat.Core
{
    /// <summary>6 hệ quy chuẩn thể thức được hỗ trợ.</summary>
    public enum StandardKind
    {
        HanhChinh_ND30 = 0,   // Nghị định 30/2020/NĐ-CP
        QuanDoi_TT68 = 1,     // Thông tư 68/2023/TT-BQP
        Dang_HD05 = 2,        // Hướng dẫn 05-HD/VPTW
        Doan_HD29 = 3,        // Hướng dẫn 29-HD/TWĐTN-VP
        PhapDien_ND250 = 4,   // Pháp điển & hợp nhất văn bản QPPL
        GiaoDuc_CV5512 = 5,   // Kế hoạch bài dạy - CV 5512/BGDĐT-GDTrH
    }

    /// <summary>Loại thành phần thể thức của một đoạn văn.</summary>
    public enum ElementKind
    {
        Empty,
        Body,
        QuocHieu,
        TieuNgu,
        TieuDeDang,        // ĐẢNG CỘNG SẢN VIỆT NAM
        TieuDeDoan,        // ĐOÀN TNCS HỒ CHÍ MINH
        CoQuanChuQuan,
        CoQuanBanHanh,
        DauSao,            // * hoặc *** dưới tên cơ quan Đảng/Đoàn
        SoKyHieu,
        TrichYeuCongVan,   // V/v ...
        DiaDanhNgayThang,
        TenLoai,
        TrichYeu,
        KinhGui,
        CanCu,             // Căn cứ ...
        Phan,              // Phần thứ nhất / Phần I
        Chuong,            // Chương I
        Muc,               // Mục 1
        TenPhanChuong,     // tên phần/chương/mục (in hoa)
        Dieu,              // Điều 1.
        MucLaMa,           // I. II.
        MucSo,             // 1. (đề mục)
        MucChuCai,         // a) b)
        TableText,
        ChucVuKy,          // TM. / KT. / chức vụ
        HoTenKy,
        NoiNhanTitle,
        NoiNhanItem,
        ThongTinDau,       // dòng thông tin trước tên loại (KHBD: Trường, Tổ, Họ tên giáo viên)
    }

    public enum Align { Left = 0, Center = 1, Right = 2, Justify = 3 }

    /// <summary>Ảnh chụp một đoạn văn bản (độc lập với Word).</summary>
    public class ParagraphInfo
    {
        public int Index;
        public string Text = "";
        public bool InTable;
        public int TableIndex;     // 1-based, 0 nếu không trong bảng
        public int CellColumn;     // 1-based
        public int CellRow;

        // Định dạng hiện tại (dùng cho rà soát). null/0/-1 = hỗn hợp hoặc không rõ.
        public string FontName;
        public float FontSize;
        public int Bold = -1;      // 1 đậm, 0 không, -1 hỗn hợp
        public int Italic = -1;
        public Align? Alignment;
        public float FirstIndentPt;

        public ParagraphInfo() { }
        public ParagraphInfo(string text) { Text = text ?? ""; }
        public override string ToString() => $"[{Index}] {Text}";
    }

    public class PageInfo
    {
        public float WidthMm, HeightMm;
        public float TopMm, BottomMm, LeftMm, RightMm;
        public bool DifferentFirstPage;
        public bool HasPageNumberInHeader;
        public bool HasPageNumberInFooter;
    }

    public enum Severity { Info, Warning, Error }

    public class Issue
    {
        public string Code;          // ví dụ ND30-03, DTN-07
        public int ParagraphIndex;   // -1 = cấp văn bản
        public string Message;
        public Severity Severity = Severity.Warning;
        public bool AutoFixable = true;
        public override string ToString() => $"{Code} (đoạn {ParagraphIndex}): {Message}";
    }

    /// <summary>Định dạng đích cho một thành phần.</summary>
    public class ElementStyle
    {
        public float Size;
        public float MinSize, MaxSize;     // khoảng cho phép khi rà soát
        public bool Bold;
        public bool Italic;
        public bool Underline;
        public bool UpperCase;
        public Align Align = Align.Justify;
        public float FirstIndentCm;
        public float SpaceBeforePt;
        public float SpaceAfterPt;
        public bool KeepLineSpacingSingle; // khối đầu / chữ ký dùng dòng đơn

        public ElementStyle Clone() => (ElementStyle)MemberwiseClone();
    }

    public class ClassifiedDocument
    {
        public List<ParagraphInfo> Paragraphs = new List<ParagraphInfo>();
        public List<ElementKind> Kinds = new List<ElementKind>();
        public StandardKind Standard;

        public int FirstIndexOf(ElementKind k) => Kinds.IndexOf(k);
        public bool Has(ElementKind k) => Kinds.Contains(k);
    }
}
