using System;
using System.Collections.Generic;
using System.Linq;

namespace VFormat.Core
{
    public class DocTemplate
    {
        public string Id;
        public string Name;
        public string Group;          // Hành chính / Quân đội / Đảng / Đoàn / Hợp nhất / Giáo dục
        public StandardKind Standard;
        public List<string> Lines = new List<string>();
        public override string ToString() => $"{Group} › {Name}";
    }

    /// <summary>
    /// Thư viện biểu mẫu dựng sẵn. Mỗi mẫu là chuỗi đoạn văn thô; khi chèn, V-Format chạy chuẩn hóa
    /// để tự dàn khối đầu, khối chữ ký thành bảng 2 cột và áp định dạng theo quy chuẩn của mẫu.
    /// </summary>
    public static class TemplateLibrary
    {
        public static string Today(DateTime? d = null)
        {
            var x = d ?? DateTime.Today;
            return TextNormalizer.FixNgayThang($"ngày {x.Day} tháng {x.Month} năm {x.Year}");
        }

        static readonly string Blank = "";

        static DocTemplate HanhChinh(string id, string name, string chuQuan, string coQuan, string kyHieu, string tenLoai,
            string trichYeu, string[] body, string chucVu, string[] noiNhan, StandardKind std = StandardKind.HanhChinh_ND30,
            string group = "Hành chính", string ten = "Nguyễn Văn A")
        {
            var t = new DocTemplate { Id = id, Name = name, Group = group, Standard = std };
            var L = t.Lines;
            if (chuQuan != null) L.Add(chuQuan);
            L.Add(coQuan);
            L.Add("Số:      /" + kyHieu);
            if (tenLoai == null) L.Add("V/v " + trichYeu);
            L.Add(TextNormalizer.QuocHieuText);
            L.Add(TextNormalizer.TieuNguText);
            L.Add("Hà Nội, " + Today());
            if (tenLoai != null) { L.Add(tenLoai); if (trichYeu != null) L.Add(trichYeu); }
            L.Add(Blank);
            L.AddRange(body);
            L.Add(Blank);
            L.Add("Nơi nhận:");
            L.AddRange(noiNhan);
            L.AddRange(chucVu.Split('|'));
            L.Add(Blank); L.Add(Blank); L.Add(Blank);
            L.Add(ten);
            return t;
        }

        static DocTemplate DangDoan(string id, string name, bool doan, string capTren, string coQuan, string kyHieu,
            string tenLoai, string trichYeu, string[] body, string chucVu, string[] noiNhan)
        {
            var t = new DocTemplate { Id = id, Name = name, Group = doan ? "Đoàn" : "Đảng",
                Standard = doan ? StandardKind.Doan_HD29 : StandardKind.Dang_HD05 };
            var L = t.Lines;
            if (capTren != null) L.Add(capTren);
            L.Add(coQuan);
            L.Add(doan ? "***" : "*");
            L.Add((doan ? "Số: " : "Số ") + "  -" + kyHieu);
            if (tenLoai == null && trichYeu != null) L.Add(trichYeu);   // công văn: trích yếu dưới số, ký hiệu
            L.Add(doan ? "ĐOÀN TNCS HỒ CHÍ MINH" : "ĐẢNG CỘNG SẢN VIỆT NAM");
            L.Add("Hà Nội, " + Today());
            if (tenLoai != null)
            {
                L.Add(tenLoai);
                if (trichYeu != null) L.Add(trichYeu);
                L.Add(doan ? "-------" : "-----");
            }
            L.Add(Blank);
            L.AddRange(body);
            L.Add(Blank);
            L.Add("Nơi nhận:");
            L.AddRange(noiNhan);
            L.AddRange(chucVu.Split('|'));
            L.Add(Blank); L.Add(Blank); L.Add(Blank);
            L.Add("Nguyễn Văn A");
            return t;
        }

        static readonly string[] NNHanhChinh = { "- Như Điều 3;", "- ...;", "- Lưu: VT, ...." };
        static readonly string[] NNCongVan = { "- Như trên;", "- ...;", "- Lưu: VT, ...." };

        static List<DocTemplate> _all;
        public static IReadOnlyList<DocTemplate> All => _all ?? (_all = Build());

        public static DocTemplate Find(string id) => All.FirstOrDefault(t => t.Id == id);

        static List<DocTemplate> Build()
        {
            var list = new List<DocTemplate>();
            string P = "Nội dung ...";

            // ===== Hành chính (NĐ 30) =====
            list.Add(HanhChinh("hc-qd", "Quyết định (cá biệt)", "ỦY BAN NHÂN DÂN TỈNH ...", "SỞ ...", "QĐ-...", "QUYẾT ĐỊNH",
                "Về việc ...",
                new[] { "GIÁM ĐỐC SỞ ...", "Căn cứ ...;", "Căn cứ ...;", "Theo đề nghị của ...", "QUYẾT ĐỊNH:",
                        "Điều 1. " + P, "Điều 2. " + P, "Điều 3. Chánh Văn phòng, ... chịu trách nhiệm thi hành Quyết định này./." },
                "GIÁM ĐỐC", NNHanhChinh));
            list.Add(HanhChinh("hc-cv", "Công văn", "ỦY BAN NHÂN DÂN TỈNH ...", "SỞ ...", "...-...", null, "...",
                new[] { "Kính gửi: ...", P, "Trân trọng./." }, "KT. GIÁM ĐỐC|PHÓ GIÁM ĐỐC", NNCongVan));
            list.Add(HanhChinh("hc-tt", "Tờ trình", "ỦY BAN NHÂN DÂN TỈNH ...", "SỞ ...", "TTr-...", "TỜ TRÌNH", "Về việc ...",
                new[] { "Kính gửi: ...", "I. SỰ CẦN THIẾT", P, "II. NỘI DUNG ĐỀ XUẤT", P, "III. KIẾN NGHỊ", P, "Kính trình ... xem xét, quyết định./." },
                "GIÁM ĐỐC", NNCongVan));
            list.Add(HanhChinh("hc-bc", "Báo cáo", "ỦY BAN NHÂN DÂN TỈNH ...", "SỞ ...", "BC-...", "BÁO CÁO", "Kết quả ...",
                new[] { "I. KẾT QUẢ THỰC HIỆN", "1. Công tác ...", P, "2. Công tác ...", P, "II. TỒN TẠI, HẠN CHẾ", P,
                        "III. PHƯƠNG HƯỚNG, NHIỆM VỤ", P, "IV. ĐỀ XUẤT, KIẾN NGHỊ", P },
                "GIÁM ĐỐC", NNCongVan));
            list.Add(HanhChinh("hc-kh", "Kế hoạch", "ỦY BAN NHÂN DÂN TỈNH ...", "SỞ ...", "KH-...", "KẾ HOẠCH", "...",
                new[] { "I. MỤC ĐÍCH, YÊU CẦU", "1. Mục đích", P, "2. Yêu cầu", P, "II. NỘI DUNG", P, "III. TỔ CHỨC THỰC HIỆN", P },
                "GIÁM ĐỐC", NNCongVan));
            list.Add(HanhChinh("hc-tb", "Thông báo", "ỦY BAN NHÂN DÂN TỈNH ...", "SỞ ...", "TB-...", "THÔNG BÁO", "...",
                new[] { P }, "GIÁM ĐỐC", NNCongVan));
            list.Add(HanhChinh("hc-gm", "Giấy mời", "ỦY BAN NHÂN DÂN TỈNH ...", "SỞ ...", "GM-...", "GIẤY MỜI", "Dự ...",
                new[] { "Sở ... trân trọng kính mời: ...", "Thời gian: ...", "Địa điểm: ...", "Nội dung: ..." },
                "GIÁM ĐỐC", NNCongVan));
            list.Add(HanhChinh("hc-bb", "Biên bản", "ỦY BAN NHÂN DÂN TỈNH ...", "SỞ ...", "BB-...", "BIÊN BẢN", "...",
                new[] { "Thời gian bắt đầu: ...", "Địa điểm: ...", "Thành phần tham dự: ...", "Chủ trì: ...", "Thư ký: ...",
                        "Nội dung: ...", "Cuộc họp kết thúc vào ... giờ ..., ngày ... tháng ... năm ..." },
                "CHỦ TỌA", NNCongVan));
            list.Add(HanhChinh("hc-gioithieu", "Giấy giới thiệu", null, "SỞ ...", "GGT-...", "GIẤY GIỚI THIỆU", null,
                new[] { "Sở ... trân trọng giới thiệu: ...", "Chức vụ: ...", "Được cử đến: ...", "Về việc: ...",
                        "Đề nghị Quý cơ quan tạo điều kiện để đồng chí có tên ở trên hoàn thành nhiệm vụ.", "Giấy này có giá trị đến hết ngày ..." },
                "GIÁM ĐỐC", NNCongVan));
            list.Add(HanhChinh("hc-uyquyen", "Giấy ủy quyền", null, "CÔNG TY ...", "GUQ-...", "GIẤY ỦY QUYỀN", null,
                new[] { "Căn cứ ...", "Tôi (người ủy quyền): ...", "Ủy quyền cho: ...", "Nội dung ủy quyền: ...", "Thời hạn ủy quyền: ..." },
                "NGƯỜI ỦY QUYỀN", NNCongVan));

            // ===== Quân đội (TT 68) =====
            list.Add(HanhChinh("qd-cv", "Công văn (Quân đội)", "BỘ CHỈ HUY QUÂN SỰ TỈNH ...", "PHÒNG THAM MƯU", "...-TM", null, "...",
                new[] { "Kính gửi: ...", P }, "TRƯỞNG PHÒNG", NNCongVan, StandardKind.QuanDoi_TT68, "Quân đội", "Đại tá Nguyễn Văn A"));
            list.Add(HanhChinh("qd-kh", "Kế hoạch (Quân đội)", "QUÂN KHU ...", "BỘ CHỈ HUY QUÂN SỰ TỈNH ...", "KH-BCH", "KẾ HOẠCH", "...",
                new[] { "I. MỤC ĐÍCH, YÊU CẦU", P, "II. NỘI DUNG", P, "III. TỔ CHỨC THỰC HIỆN", P },
                "CHỈ HUY TRƯỞNG", NNCongVan, StandardKind.QuanDoi_TT68, "Quân đội", "Đại tá Nguyễn Văn A"));
            list.Add(HanhChinh("qd-bc", "Báo cáo (Quân đội)", "SƯ ĐOÀN ...", "TRUNG ĐOÀN ...", "BC-TrĐ", "BÁO CÁO", "...",
                new[] { "I. TÌNH HÌNH", P, "II. KẾT QUẢ", P, "III. ĐỀ XUẤT", P },
                "TRUNG ĐOÀN TRƯỞNG", NNCongVan, StandardKind.QuanDoi_TT68, "Quân đội", "Thượng tá Nguyễn Văn A"));
            list.Add(HanhChinh("qd-qd", "Quyết định (Quân đội)", "QUÂN KHU ...", "BỘ CHỈ HUY QUÂN SỰ TỈNH ...", "QĐ-BCH", "QUYẾT ĐỊNH", "Về việc ...",
                new[] { "CHỈ HUY TRƯỞNG BỘ CHỈ HUY QUÂN SỰ TỈNH ...", "Căn cứ ...;", "QUYẾT ĐỊNH:", "Điều 1. " + P, "Điều 2. " + P },
                "CHỈ HUY TRƯỞNG", NNHanhChinh, StandardKind.QuanDoi_TT68, "Quân đội", "Đại tá Nguyễn Văn A"));

            // ===== Công an (08/VBHN-BCA) =====
            const StandardKind CA = StandardKind.CongAn_BCA08;
            list.Add(HanhChinh("ca-cv", "Công văn (Công an)", "CÔNG AN TỈNH ...", "PHÒNG THAM MƯU", "...-PV01", null, "...",
                new[] { "Kính gửi: ...", P, "Trân trọng./." }, "KT. TRƯỞNG PHÒNG|PHÓ TRƯỞNG PHÒNG", NNCongVan, CA, "Công an", "Thượng tá Nguyễn Văn A"));
            list.Add(HanhChinh("ca-kh", "Kế hoạch (Công an)", "BỘ CÔNG AN", "CÔNG AN TỈNH ...", "KH-CAT-PV01", "KẾ HOẠCH", "...",
                new[] { "I. MỤC ĐÍCH, YÊU CẦU", P, "II. NỘI DUNG", P, "III. TỔ CHỨC THỰC HIỆN", P },
                "GIÁM ĐỐC", NNCongVan, CA, "Công an", "Đại tá Nguyễn Văn A"));
            list.Add(HanhChinh("ca-bc", "Báo cáo (Công an)", "BỘ CÔNG AN", "CÔNG AN TỈNH ...", "BC-CAT-PV01", "BÁO CÁO", "...",
                new[] { "I. TÌNH HÌNH", P, "II. KẾT QUẢ CÔNG TÁC", P, "III. PHƯƠNG HƯỚNG, NHIỆM VỤ", P },
                "GIÁM ĐỐC", NNCongVan, CA, "Công an", "Đại tá Nguyễn Văn A"));
            list.Add(HanhChinh("ca-qd", "Quyết định (Công an)", "BỘ CÔNG AN", "CÔNG AN TỈNH ...", "QĐ-CAT-PX01", "QUYẾT ĐỊNH", "Về việc ...",
                new[] { "GIÁM ĐỐC CÔNG AN TỈNH ...", "Căn cứ ...;", "Theo đề nghị của ...", "QUYẾT ĐỊNH:", "Điều 1. " + P, "Điều 2. " + P,
                        "Điều 3. ... chịu trách nhiệm thi hành Quyết định này./." },
                "GIÁM ĐỐC", NNHanhChinh, CA, "Công an", "Đại tá Nguyễn Văn A"));

            // ===== Đảng (HD 05) =====
            list.Add(DangDoan("d-nq", "Nghị quyết", false, "ĐẢNG BỘ TỈNH ...", "HUYỆN ỦY ...", "NQ/HU", "NGHỊ QUYẾT", "về ...",
                new[] { "I. TÌNH HÌNH", P, "II. MỤC TIÊU, NHIỆM VỤ", P, "III. TỔ CHỨC THỰC HIỆN", P },
                "T/M BAN CHẤP HÀNH|BÍ THƯ", new[] { "- Tỉnh ủy (báo cáo);", "- Lưu Văn phòng Huyện ủy." }));
            list.Add(DangDoan("d-bc", "Báo cáo (Đảng)", false, "ĐẢNG BỘ TỈNH ...", "HUYỆN ỦY ...", "BC/HU", "BÁO CÁO", "kết quả ...",
                new[] { "I. KẾT QUẢ", P, "II. HẠN CHẾ", P, "III. NHIỆM VỤ", P },
                "T/M BAN THƯỜNG VỤ|BÍ THƯ", new[] { "- Tỉnh ủy (báo cáo);", "- Lưu Văn phòng Huyện ủy." }));
            list.Add(DangDoan("d-tb", "Thông báo (Đảng)", false, "ĐẢNG BỘ TỈNH ...", "HUYỆN ỦY ...", "TB/HU", "THÔNG BÁO", "...",
                new[] { P }, "T/M BAN THƯỜNG VỤ|BÍ THƯ", new[] { "- ...;", "- Lưu Văn phòng Huyện ủy." }));
            list.Add(DangDoan("d-qd", "Quyết định (Đảng)", false, "ĐẢNG BỘ TỈNH ...", "HUYỆN ỦY ...", "QĐ/HU", "QUYẾT ĐỊNH", "về ...",
                new[] { "- Căn cứ Điều lệ Đảng;", "- Căn cứ Quy chế làm việc của Huyện ủy.", "BAN THƯỜNG VỤ HUYỆN ỦY QUYẾT ĐỊNH", "Điều 1. " + P, "Điều 2. " + P },
                "T/M BAN THƯỜNG VỤ|BÍ THƯ", new[] { "- Như Điều 2;", "- Lưu Văn phòng Huyện ủy." }));

            // ===== Đoàn (HD 29) =====
            list.Add(DangDoan("dtn-kh", "Kế hoạch (Đoàn)", true, "BCH ĐOÀN TỈNH ...", "BCH ĐOÀN HUYỆN ...", "KH/HĐTN", "KẾ HOẠCH", "...",
                new[] { "I. MỤC ĐÍCH, YÊU CẦU", P, "II. NỘI DUNG", P, "III. TỔ CHỨC THỰC HIỆN", P },
                "TM. BAN THƯỜNG VỤ HUYỆN ĐOÀN|BÍ THƯ", new[] { "- Tỉnh Đoàn (b/c);", "- Lưu VP." }));
            list.Add(DangDoan("dtn-bc", "Báo cáo (Đoàn)", true, "BCH ĐOÀN TỈNH ...", "BCH ĐOÀN HUYỆN ...", "BC/HĐTN", "BÁO CÁO", "...",
                new[] { "I. KẾT QUẢ", P, "II. TỒN TẠI", P, "III. PHƯƠNG HƯỚNG", P },
                "TM. BAN THƯỜNG VỤ HUYỆN ĐOÀN|BÍ THƯ", new[] { "- Tỉnh Đoàn (b/c);", "- Lưu VP." }));
            list.Add(DangDoan("dtn-cv", "Công văn (Đoàn)", true, "BCH ĐOÀN TỈNH ...", "BCH ĐOÀN HUYỆN ...", "CV/HĐTN", null, "V/v ...",
                new[] { "Kính gửi: ...", P }, "TM. BAN THƯỜNG VỤ HUYỆN ĐOÀN|PHÓ BÍ THƯ", new[] { "- Như kính gửi;", "- Lưu VP." }));

            // ===== Văn bản hợp nhất (NĐ 250/2026/NĐ-CP, Phụ lục: sơ đồ mục I.2, ví dụ 2, phần ký xác thực mục XIII) =====
            list.Add(new DocTemplate
            {
                Id = "pd-hn", Name = "Văn bản hợp nhất", Group = "Hợp nhất", Standard = StandardKind.PhapDien_ND250,
                Lines =
                {
                    "VĂN BẢN HỢP NHẤT", "Số:      /" + DateTime.Today.Year + "/VBHN-TT-...", TextNormalizer.QuocHieuText, TextNormalizer.TieuNguText,
                    "Hà Nội, " + Today(),
                    "THÔNG TƯ", "Quy định ...", Blank,
                    "Thông tư số .../.../TT-... ngày ... tháng ... năm ... của ..., có hiệu lực kể từ ngày ... tháng ... năm ..., được sửa đổi, bổ sung bởi:",
                    "Thông tư số .../.../TT-... ngày ... tháng ... năm ... của ... sửa đổi, bổ sung một số điều của ..., có hiệu lực kể từ ngày ... tháng ... năm ....",
                    "Căn cứ ...;",
                    "Chương I", "QUY ĐỊNH CHUNG", "Điều 1. Phạm vi điều chỉnh", P, "Điều 2. Đối tượng áp dụng", P,
                    "Chương II", "QUY ĐỊNH CỤ THỂ", "Điều 3. ...", "1. " + P + ".", "a) " + P + ";", "b) " + P + ".",
                    Blank, "Nơi nhận:", "- Văn phòng Chính phủ (để đăng Công báo);", "- Cơ quan có thẩm quyền (để đăng trên Cơ sở dữ liệu quốc gia về pháp luật);",
                    "- Lưu: VT, ....", "BỘ ...", "XÁC THỰC VĂN BẢN HỢP NHẤT", "Hà Nội, " + Today(), "BỘ TRƯỞNG", "(Chữ ký và dấu)", Blank, Blank, "Nguyễn Văn A",
                },
            });

            // ===== Giáo dục (CV 5512) =====
            list.Add(new DocTemplate
            {
                Id = "gd-khbd", Name = "Kế hoạch bài dạy (Phụ lục IV)", Group = "Giáo dục", Standard = StandardKind.GiaoDuc_CV5512,
                Lines =
                {
                    "Trường: ...", "Tổ: ...", "Họ và tên giáo viên: ...",
                    "KẾ HOẠCH BÀI DẠY", "TÊN BÀI DẠY: ...", "Môn học/Hoạt động giáo dục: ...; lớp: ...", "Thời gian thực hiện: (số tiết)",
                    "I. Mục tiêu", "1. Về kiến thức: " + P, "2. Về năng lực: " + P, "3. Về phẩm chất: " + P,
                    "II. Thiết bị dạy học và học liệu", P,
                    "III. Tiến trình dạy học",
                    "1. Hoạt động 1: Xác định vấn đề/nhiệm vụ học tập/Mở đầu", "a) Mục tiêu", P, "b) Nội dung", P, "c) Sản phẩm", P, "d) Tổ chức thực hiện", P,
                    "2. Hoạt động 2: Hình thành kiến thức mới/giải quyết vấn đề/thực thi nhiệm vụ đặt ra từ Hoạt động 1", "a) Mục tiêu", P, "b) Nội dung", P, "c) Sản phẩm", P, "d) Tổ chức thực hiện", P,
                    "3. Hoạt động 3: Luyện tập", "a) Mục tiêu", P, "b) Nội dung", P, "c) Sản phẩm", P, "d) Tổ chức thực hiện", P,
                    "4. Hoạt động 4: Vận dụng", "a) Mục tiêu", P, "b) Nội dung", P, "c) Sản phẩm", P, "d) Tổ chức thực hiện", P,
                },
            });
            return list;
        }
    }
}
