using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

namespace VFormat.Core
{
    /// <summary>
    /// Nhận diện thành phần thể thức cho từng đoạn: khối đầu (Quốc hiệu, cơ quan, số ký hiệu, ngày tháng),
    /// tên loại + trích yếu, đề mục nội dung, khối chữ ký và nơi nhận.
    /// </summary>
    public static class Classifier
    {
        static readonly string[] DocTypes =
        {
            "QUYET DINH", "NGHI QUYET", "NGHI DINH", "THONG TU", "CHI THI", "QUY CHE", "QUY DINH", "THONG BAO",
            "HUONG DAN", "CHUONG TRINH", "KE HOACH", "PHUONG AN", "DE AN", "DU AN", "BAO CAO", "BIEN BAN",
            "TO TRINH", "HOP DONG", "CONG DIEN", "BAN GHI NHO", "BAN THOA THUAN", "GIAY UY QUYEN", "GIAY MOI",
            "GIAY GIOI THIEU", "GIAY NGHI PHEP", "PHIEU GUI", "PHIEU CHUYEN", "PHIEU BAO", "THU CONG", "KET LUAN",
            "CONG VAN", "VAN BAN HOP NHAT", "LENH", "MENH LENH", "CHI LENH", "DIEN", "THU", "QUY TRINH",
            "KE HOACH BAI DAY", "GIAO AN", "DIEU LE", "TRICH YEU", "DANH SACH", "BAN CAM KET", "DON",
        };

        static readonly Regex RxSo = new Regex(@"^Số\s*:?\s*[\d\.…_ ]*[/\-]", RegexOptions.IgnoreCase);
        static readonly Regex RxNgayThang = new Regex(@"ngày\s*[\d\.…_ ]*\s*tháng\s*[\d\.…_ ]*\s*năm\s*[\d\.…_ ]*", RegexOptions.IgnoreCase);
        static readonly Regex RxPhan = new Regex(@"^Phần\s+(thứ\s+\S+|[IVXLC]+\b|\d+\b)", RegexOptions.IgnoreCase);
        static readonly Regex RxChuong = new Regex(@"^Chương\s+([IVXLC]+|\d+)\b", RegexOptions.IgnoreCase);
        static readonly Regex RxMuc = new Regex(@"^Mục\s+\d+\b");
        static readonly Regex RxDieu = new Regex(@"^Điều\s+\d+[a-z]?\s*[\.:]");
        static readonly Regex RxLaMa = new Regex(@"^[IVXLC]+\s*[\.\)\-]\s+\S");
        // "1.", "1.2)", "2.Nội dung" — không nhận số tiền "1.000.000đ" (sau dấu chấm phải là khoảng trắng hoặc chữ)
        static readonly Regex RxSoMuc = new Regex(@"^\d+(\.\d+)*\s*[\.\)](\s+\S|\p{L})");
        static readonly Regex RxHoatDong = new Regex(@"^Hoạt động\s+\d+", RegexOptions.IgnoreCase);
        static readonly Regex RxChuCai = new Regex(@"^[a-zđ]\s*[\)\.]\s+\S");
        static readonly Regex RxDecor = new Regex(@"^[\-_–—=\*\.]{2,}$|^\*$");
        static readonly Regex RxChucVuPrefix = new Regex(@"^(T\.?M\.?|T/M|K\.?T\.?|K/T|Q\.|T\.?L\.?|T/L|TUQ\.?|T/UQ|THỪA LỆNH|THỪA ỦY QUYỀN|THỪA UỶ QUYỀN)\s", RegexOptions.IgnoreCase);

        /// <summary>Dòng quyền hạn ký ("TM. ...", "KT. ...", "T/M ...") — khác dòng chức vụ người ký ("PHÓ GIÁM ĐỐC").</summary>
        public static bool HasSignAuthority(string s) => RxChucVuPrefix.IsMatch(TextUtil.Clean(s));
        static readonly string[] ChucVuKeys =
        {
            "CHU TICH", "GIAM DOC", "BO TRUONG", "THU TRUONG", "BI THU", "TRUONG", "PHO ", "CHI HUY", "TU LENH",
            "CHINH UY", "HIEU TRUONG", "TONG ", "THU TUONG", "NGUOI ", "CHANH ", "VIEN TRUONG", "DAI DIEN",
            "UY VIEN", "THU KY", "CHU NHIEM", "CHU TOA", "GIAO VIEN", "XAC NHAN", "KE TOAN", "TAP THE",
        };

        /// <summary>Ghi chú biểu mẫu ở đầu (cuối) trang: "Mẫu BL2", "BH theo TT số 44/2025/TT-BCA", "Ngày 18/6/2025", "VB21".</summary>
        static readonly Regex RxGhiChuMau = new Regex(
            @"^(Mẫu( số)?\s*[:.]?\s*\S{1,15}|(BH|Ban hành)( kèm)? theo\b.{0,100}|Ngày \d{1,2}/\d{1,2}/\d{4}|[A-Z]{1,4}\d{1,4}[a-z]?)$");
        // "(Ký, ghi rõ họ tên)"; biểu mẫu có thể dính mã mẫu phía sau: "(Ký, ghi rõ họ tên và đóng dấu) VB21"
        static readonly Regex RxGhiChuKy = new Regex(@"^\(.*\)(\s+[A-Z]{1,4}\d{1,4}[a-z]?)?$");
        // dòng thời gian của chữ ký số ("11:46:03 +07:00") nằm giữa chức vụ và họ tên người ký
        static readonly Regex RxSignTime = new Regex(@"^\d{1,2}:\d{2}(:\d{2})?\s*([+\-]\d{2}:?\d{2})?$");

        /// <summary>Loại của phần phải trong đoạn khối đầu gõ hai cột bằng tab; null nếu không phải.</summary>
        static ElementKind? HeaderRightKind(string s)
        {
            var ky = TextUtil.Key(s);
            if (ky == "CONG HOA XA HOI CHU NGHIA VIET NAM") return ElementKind.QuocHieu;
            if (ky.StartsWith("DOC LAP") && ky.Contains("TU DO") && ky.Length < 45) return ElementKind.TieuNgu;
            if (ky == "DANG CONG SAN VIET NAM") return ElementKind.TieuDeDang;
            if (Regex.IsMatch(ky, @"^DOAN (TNCS|THANH NIEN CONG SAN) HO CHI MINH$")) return ElementKind.TieuDeDoan;
            if (RxNgayThang.IsMatch(s) && s.Length < 90 && s.IndexOf(',') > 0) return ElementKind.DiaDanhNgayThang;
            return null;
        }

        /// <summary>Loại của phần trái (số ký hiệu, trích yếu công văn, dấu phân cách, tên cơ quan); null nếu không nhận được.</summary>
        static ElementKind? HeaderLeftKind(string s)
        {
            if (RxSo.IsMatch(s) && s.Length < 60) return ElementKind.SoKyHieu;
            if (TextUtil.Key(s).StartsWith("V/V")) return ElementKind.TrichYeuCongVan;
            if (RxDecor.IsMatch(s)) return ElementKind.DauSao;
            if (TextUtil.IsAllUpper(s)) return ElementKind.CoQuanBanHanh;   // chủ quản / ban hành xác định ở bước 2
            return null;
        }

        /// <summary>Số thứ tự tự động của Word ("II.", "1.", "a)") — không phải dấu gạch đầu dòng.</summary>
        static bool IsNumberPrefix(string prefix) =>
            !string.IsNullOrEmpty(prefix) && Regex.IsMatch(prefix.Trim(), @"^([IVXLCivxlc]+|\d+(\.\d+)*|[a-zđA-Z])\s*[\.\)\-:]?$");

        public static ClassifiedDocument Classify(IList<ParagraphInfo> paragraphs, StandardKind standard)
        {
            var doc = new ClassifiedDocument { Standard = standard };
            doc.Paragraphs.AddRange(paragraphs);
            int n = paragraphs.Count;
            var t = paragraphs.Select(p => TextUtil.Clean(p.Text)).ToArray();
            var key = t.Select(TextUtil.Key).ToArray();
            var k = new ElementKind?[n];
            // chữ dùng để nhận đề mục: kèm số thứ tự tự động của Word ("II." không có trong Range.Text)
            var num = t.Select((s, i) => IsNumberPrefix(paragraphs[i].ListPrefix) && s.Length > 0
                ? paragraphs[i].ListPrefix.Trim() + " " + s : s).ToArray();

            // 1. Mẫu độc lập ngữ cảnh
            int firstHeader = -1;
            for (int i = 0; i < n; i++)
            {
                if (t[i].Length == 0) { k[i] = ElementKind.Empty; continue; }
                var s = t[i]; var ky = key[i];
                // khối đầu gõ hai cột bằng tab: "CỤC CẢNH SÁT ⇥ CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM"
                var parts = i < 20 ? TextUtil.TabParts(paragraphs[i].Text) : new string[0];
                if (parts.Length >= 2)
                {
                    var right = parts[parts.Length - 1];
                    var left = string.Join(" ", parts.Take(parts.Length - 1));
                    var rk = HeaderRightKind(right);
                    var lk = HeaderLeftKind(left);
                    if (rk.HasValue && lk.HasValue)
                    {
                        k[i] = rk.Value;
                        doc.Splits[i] = new HeaderSplit { Left = left, Right = right, LeftKind = lk.Value };
                        if (firstHeader < 0) firstHeader = i;
                        continue;
                    }
                }
                bool topOrBottom = (firstHeader < 0 && i < 8) || i >= n - 3;
                bool small = firstHeader < 0 && i < 8 && paragraphs[i].FontSize > 0 && paragraphs[i].FontSize <= 10;
                if (topOrBottom && s.Length < 120 && (RxGhiChuMau.IsMatch(s) || small))
                    k[i] = ElementKind.GhiChuMau;
                else if (ky == "CONG HOA XA HOI CHU NGHIA VIET NAM") k[i] = ElementKind.QuocHieu;
                else if (ky.StartsWith("DOC LAP") && ky.Contains("TU DO") && ky.Length < 45) k[i] = ElementKind.TieuNgu;
                else if (ky == "DANG CONG SAN VIET NAM") k[i] = ElementKind.TieuDeDang;
                else if (Regex.IsMatch(ky, @"^DOAN (TNCS|THANH NIEN CONG SAN) HO CHI MINH$")) k[i] = ElementKind.TieuDeDoan;
                else if (RxDecor.IsMatch(s)) k[i] = ElementKind.DauSao;
                else if (RxSo.IsMatch(s) && s.Length < 60) k[i] = ElementKind.SoKyHieu;
                else if (ky.StartsWith("V/V") || ky.StartsWith("VE VIEC ") && i < 15 && paragraphs[i].InTable) k[i] = ElementKind.TrichYeuCongVan;
                else if (ky == "XAC THUC VAN BAN HOP NHAT") k[i] = ElementKind.XacThuc;
                else if (RxNgayThang.IsMatch(s) && s.Length < 90 && s.IndexOf(',') > 0
                         && (i < 25 || paragraphs[i].InTable || PrevNonEmpty(k, i) == ElementKind.XacThuc)) k[i] = ElementKind.DiaDanhNgayThang;
                else if (ky.StartsWith("KINH GUI")) k[i] = ElementKind.KinhGui;
                else if (Regex.IsMatch(ky, @"^NOI NHAN\s*:?$")) k[i] = ElementKind.NoiNhanTitle;
                if (firstHeader < 0 && k[i].HasValue && k[i] != ElementKind.GhiChuMau) firstHeader = i;
            }

            // tên cơ quan gõ tab bị xuống dòng: "CỤC CẢNH SÁT GIAO ⇥ CỘNG HÒA..." / "THÔNG ⇥ Độc lập..." → nối vào dòng trên
            foreach (var i in doc.Splits.Keys.OrderBy(x => x).ToList())
            {
                var sp = doc.Splits[i];
                if (sp.LeftKind != ElementKind.CoQuanBanHanh || TextUtil.WordCount(sp.Left) != 1) continue;
                if (doc.Splits.TryGetValue(i - 1, out var prev) && prev.LeftKind == ElementKind.CoQuanBanHanh)
                {
                    prev.Left += " " + sp.Left;
                    sp.LeftKind = ElementKind.Empty;
                }
            }

            // 2. Vùng khối đầu
            var headerKinds = new[] { ElementKind.QuocHieu, ElementKind.TieuNgu, ElementKind.TieuDeDang, ElementKind.TieuDeDoan,
                                      ElementKind.SoKyHieu, ElementKind.DiaDanhNgayThang, ElementKind.TrichYeuCongVan };
            int headerEnd = -1;
            for (int i = 0; i < n && i < 20; i++)
            {
                // văn bản ngắn: ngày tháng của phần ký xác thực / chữ ký không thuộc khối đầu
                if (k[i] == ElementKind.XacThuc || k[i] == ElementKind.NoiNhanTitle || k[i] == ElementKind.KinhGui) break;
                if (k[i].HasValue && headerKinds.Contains(k[i].Value)) headerEnd = i;
            }
            var headerTables = new HashSet<int>(Enumerable.Range(0, n)
                .Where(i => paragraphs[i].InTable && k[i].HasValue && headerKinds.Contains(k[i].Value) && i <= headerEnd)
                .Select(i => paragraphs[i].TableIndex));
            if (headerEnd >= 0)
            {
                // hết bảng khối đầu thì mới hết vùng đầu
                while (headerEnd + 1 < n && paragraphs[headerEnd + 1].InTable && headerTables.Contains(paragraphs[headerEnd + 1].TableIndex))
                    headerEnd++;
                var coQuan = new List<int>();
                for (int i = 0; i <= headerEnd; i++)
                {
                    if (doc.Splits.TryGetValue(i, out var sp)) { if (sp.LeftKind == ElementKind.CoQuanBanHanh) coQuan.Add(i); continue; }
                    if (k[i].HasValue) continue;
                    if (TextUtil.IsAllUpper(t[i]) && paragraphs[i].CellColumn <= 1) coQuan.Add(i);
                    else if (i > 0 && k[i - 1] == ElementKind.TrichYeuCongVan) k[i] = ElementKind.TrichYeuCongVan;
                    else if (TextUtil.IsAllUpper(t[i])) coQuan.Add(i);
                }
                // Cơ quan đứng trước số ký hiệu / dấu sao
                for (int j = 0; j < coQuan.Count; j++)
                {
                    var kind = j == coQuan.Count - 1 ? ElementKind.CoQuanBanHanh : ElementKind.CoQuanChuQuan;
                    if (doc.Splits.TryGetValue(coQuan[j], out var sp)) sp.LeftKind = kind;
                    else k[coQuan[j]] = kind;
                }
            }

            // 3. Tên loại + trích yếu
            int tenLoai = -1, seen = 0;
            for (int i = headerEnd + 1; i < n && seen < 6; i++)
            {
                if (k[i] == ElementKind.Empty || k[i] == ElementKind.DauSao) continue;
                seen++;
                if (k[i].HasValue) break;
                if (RxChuong.IsMatch(t[i]) || RxPhan.IsMatch(t[i]) || RxDieu.IsMatch(t[i])) break;   // đã vào nội dung
                var ky = key[i];
                bool known = DocTypes.Any(d => ky == d || ky.StartsWith(d + " ") || ky.StartsWith(d + ":"));
                if (TextUtil.IsAllUpper(t[i]) && (known || TextUtil.WordCount(t[i]) <= 8) && !RxChuong.IsMatch(t[i]))
                {
                    tenLoai = i; k[i] = ElementKind.TenLoai; break;
                }
                if (known && TextUtil.WordCount(t[i]) <= 6) { tenLoai = i; k[i] = ElementKind.TenLoai; break; }
            }
            if (tenLoai >= 0)
            {
                int count = 0;
                for (int i = tenLoai + 1; i < n && count < 3; i++)
                {
                    if (k[i] == ElementKind.Empty) { if (count > 0) break; continue; }
                    if (k[i] == ElementKind.DauSao) break;
                    if (k[i].HasValue) break;
                    var s = t[i];
                    if (Regex.IsMatch(key[i], @"^(CAN CU|KINH GUI|THUC HIEN|CHUONG|DIEU \d|I\.|1\.)") || s.Length > 250 || s.EndsWith(":"))
                        break;
                    // trích yếu không kết thúc bằng dấu câu: dòng thứ hai là câu văn thì đã sang phần nội dung
                    if (count > 0 && Regex.IsMatch(s, @"[\.;!?]$") && !s.EndsWith("...")) break;
                    if (count == 0 && TextUtil.IsAllUpper(s) && DocTypes.Any(d => key[i] == d))
                    {
                        k[i] = ElementKind.TenLoai; continue;   // "VĂN BẢN HỢP NHẤT" ↵ "THÔNG TƯ"
                    }
                    // dòng in hoa sau tên loại là thẩm quyền ban hành ("GIÁM ĐỐC SỞ ..."), không phải trích yếu
                    if (TextUtil.IsAllUpper(s) && standard != StandardKind.GiaoDuc_CV5512) break;
                    k[i] = ElementKind.TrichYeu; count++;
                }
            }

            // 3b. Dòng thông tin đứng trước tên loại khi không có khối Quốc hiệu (KHBD)
            if (tenLoai > 0 && headerEnd < 0)
                for (int i = 0; i < tenLoai; i++)
                    if (!k[i].HasValue) k[i] = ElementKind.ThongTinDau;

            // 4. Nơi nhận
            int noiNhan = -1;
            for (int i = 0; i < n; i++) if (k[i] == ElementKind.NoiNhanTitle) noiNhan = i;
            if (noiNhan >= 0)
            {
                for (int i = noiNhan + 1; i < n; i++)
                {
                    if (k[i] == ElementKind.Empty) continue;
                    var s = t[i];
                    if (k[i] == null && (Regex.IsMatch(s, @"^[\-\+•–]") || TextUtil.Key(s).StartsWith("LUU") || TextUtil.Key(s).StartsWith("NHU TREN")))
                        k[i] = ElementKind.NoiNhanItem;
                    else break;
                }
            }

            // 5. Khối chữ ký: duyệt phần đuôi văn bản
            int bodyStart = System.Math.Max(headerEnd, tenLoai) + 1;
            int tailStart = System.Math.Max(bodyStart, n - 30);
            var signTables = new HashSet<int>();
            var chucVu = new List<int>();
            for (int i = tailStart; i < n; i++)
            {
                if (k[i].HasValue) continue;
                var s = t[i];
                bool prefix = RxChucVuPrefix.IsMatch(s);   // "TM. Ban Chấp hành" viết thường vẫn là quyền hạn ký
                if (!(TextUtil.IsAllUpper(s) || prefix) || TextUtil.WordCount(s) > 14 || s.EndsWith(".") && !prefix) continue;
                // đề mục "V. Ý KIẾN CỦA LÃNH ĐẠO/CHỈ HUY" không phải chức vụ người ký
                if (RxLaMa.IsMatch(num[i]) || RxSoMuc.IsMatch(num[i]) || RxChuCai.IsMatch(num[i])) continue;
                var prev = PrevNonEmpty(k, i);
                if (prev.HasValue && (prev == ElementKind.Chuong || prev == ElementKind.Phan || prev == ElementKind.Muc)) continue;
                // khối chữ ký đặt ở nửa phải trang bằng thụt lề trái (không dùng bảng)
                bool indentedRight = !paragraphs[i].InTable && paragraphs[i].LeftIndentPt >= 140;
                bool looks = prefix || indentedRight || ChucVuKeys.Any(c => key[i].Contains(c));
                bool afterBody = paragraphs[i].InTable || NextNonEmptyIsNameOrUpper(t, i) || i > n - 12;
                if (looks && afterBody)
                {
                    k[i] = ElementKind.ChucVuKy; chucVu.Add(i);
                    if (paragraphs[i].InTable) signTables.Add(paragraphs[i].TableIndex);
                }
            }
            if (noiNhan >= 0 && paragraphs[noiNhan].InTable) signTables.Add(paragraphs[noiNhan].TableIndex);
            // Phần ký xác thực văn bản hợp nhất (NĐ 250/2026): tên cơ quan của người ký xác thực ngay trên "XÁC THỰC VĂN BẢN HỢP NHẤT"
            for (int i = 0; i < n; i++)
            {
                if (k[i] != ElementKind.XacThuc) continue;
                if (paragraphs[i].InTable) signTables.Add(paragraphs[i].TableIndex);
                int j = i - 1;
                while (j >= 0 && k[j] == ElementKind.Empty) j--;
                if (j >= 0 && (k[j] == null || k[j] == ElementKind.ChucVuKy || k[j] == ElementKind.Body)
                    && (TextUtil.IsAllUpper(t[j]) || RxGhiChuKy.IsMatch(t[j])) && TextUtil.WordCount(t[j]) <= 16)
                {
                    k[j] = ElementKind.CoQuanXacThuc; chucVu.Remove(j);
                }
            }
            // Sau mỗi chức vụ: chức vụ kế tiếp ("KT. GIÁM ĐỐC" / "PHÓ GIÁM ĐỐC"), "(Ký, ghi rõ họ tên)", họ tên
            foreach (var c in chucVu)
            {
                for (int i = c + 1; i < n; i++)
                {
                    if (k[i] == ElementKind.Empty) continue;
                    if (k[i].HasValue) break;
                    if (IsPersonName(t[i])) { k[i] = ElementKind.HoTenKy; break; }
                    if (RxGhiChuKy.IsMatch(t[i])) { k[i] = ElementKind.GhiChuKy; continue; } // (Ký, ghi rõ họ tên) / (Đã ký)
                    // khung chữ ký số ("Ký bởi: ... Ngày ký: ...") và dòng giờ ký đứng giữa chức vụ và họ tên
                    if (paragraphs[i].InTable && !signTables.Contains(paragraphs[i].TableIndex) || RxSignTime.IsMatch(t[i])) continue;
                    break;
                }
            }
            // Ngày tháng đặt ngay trên chức vụ người ký (báo cáo, đơn, biểu mẫu): "Hà Nội, ngày 21 tháng 9 năm 2026"
            foreach (var c in chucVu)
            {
                int j = c - 1;
                while (j >= 0 && k[j] == ElementKind.Empty) j--;
                if (j >= 0 && k[j] == null && !paragraphs[j].InTable && RxNgayThang.IsMatch(t[j]) && t[j].Length < 90 && t[j].IndexOf(',') > 0)
                    k[j] = ElementKind.DiaDanhNgayThang;
            }
            // Đoạn còn lại trong bảng chữ ký
            for (int i = 0; i < n; i++)
            {
                if (k[i].HasValue || !paragraphs[i].InTable || !signTables.Contains(paragraphs[i].TableIndex)) continue;
                if (paragraphs[i].CellColumn <= 1) k[i] = ElementKind.NoiNhanItem;
                else if (RxGhiChuKy.IsMatch(t[i])) k[i] = ElementKind.GhiChuKy;
                else if (TextUtil.IsAllUpper(t[i])) k[i] = ElementKind.ChucVuKy;
                else if (IsPersonName(t[i])) k[i] = ElementKind.HoTenKy;
                else k[i] = ElementKind.TableText;
            }

            // 6. Cấu trúc nội dung
            bool afterStructural = false, contentStarted = false;
            for (int i = 0; i < n; i++)
            {
                if (k[i].HasValue)
                {
                    if (k[i] != ElementKind.Empty) afterStructural = false;
                    continue;
                }
                var s = num[i];
                bool listed = !string.IsNullOrEmpty(paragraphs[i].ListPrefix);
                ElementKind kind;
                if (RxPhan.IsMatch(s) && s.Length < 120) kind = ElementKind.Phan;
                else if (RxChuong.IsMatch(s) && s.Length < 160) kind = ElementKind.Chuong;
                else if (RxMuc.IsMatch(s) && s.Length < 160) kind = ElementKind.Muc;
                // tên phần/chương/mục ngay dưới dòng "Chương I" (kể cả khi viết thường: "quy định chung")
                else if (afterStructural && (TextUtil.IsAllUpper(s) || TextUtil.WordCount(s) <= 15 && !Regex.IsMatch(s, @"[\.;:,]$") && !RxDieu.IsMatch(s)))
                    kind = ElementKind.TenPhanChuong;
                else if (paragraphs[i].InTable) kind = ElementKind.TableText;
                else if (RxDieu.IsMatch(s)) kind = ElementKind.Dieu;
                // căn cứ ban hành chỉ đứng trước Điều / đề mục đầu tiên; "Căn cứ kế hoạch..." trong nội dung là lời văn
                else if (Regex.IsMatch(key[i], @"^(- ?)?CAN CU ") && i < bodyStart + 15 && !contentStarted) kind = ElementKind.CanCu;   // Đảng: "- Căn cứ ..."
                // "Theo đề nghị của...", "Xét đề nghị..." ngay sau các căn cứ: dòng cuối phần căn cứ ban hành (in nghiêng)
                else if (Regex.IsMatch(key[i], @"^(THEO|XET) (DE NGHI|DE XUAT|TO TRINH|KET QUA)") && PrevNonEmpty(k, i) == ElementKind.CanCu)
                    kind = ElementKind.CanCu;
                else if (RxLaMa.IsMatch(s) && s.Length < 200) kind = ElementKind.MucLaMa;
                else if ((RxSoMuc.IsMatch(s) || RxHoatDong.IsMatch(s)) && IsHeadingLike(s)) kind = ElementKind.MucSo;
                else if (RxChuCai.IsMatch(s) && IsHeadingLike(s) && s.Length < 90) kind = ElementKind.MucChuCai;
                else if (TextUtil.IsAllUpper(s) && TextUtil.WordCount(s) <= 12 && i >= bodyStart && !listed) kind = ElementKind.TenPhanChuong;
                else kind = ElementKind.Body;
                k[i] = kind;
                afterStructural = kind == ElementKind.Phan || kind == ElementKind.Chuong || kind == ElementKind.Muc;
                if (afterStructural || kind == ElementKind.Dieu || kind == ElementKind.MucLaMa || kind == ElementKind.MucSo || kind == ElementKind.MucChuCai)
                    contentStarted = true;
            }

            doc.Kinds.AddRange(k.Select(x => x ?? ElementKind.Body));
            return doc;
        }

        static ElementKind? PrevNonEmpty(ElementKind?[] k, int i)
        {
            for (int j = i - 1; j >= 0; j--) if (k[j] != ElementKind.Empty) return k[j];
            return null;
        }

        static bool NextNonEmptyIsNameOrUpper(string[] t, int i)
        {
            for (int j = i + 1; j < t.Length && j < i + 8; j++)
            {
                if (t[j].Length == 0) continue;
                return IsPersonName(t[j]) || TextUtil.IsAllUpper(t[j]) || t[j].StartsWith("(");
            }
            return true; // cuối văn bản
        }

        static bool IsHeadingLike(string s) =>
            s.Length <= 160 && !Regex.IsMatch(s, @"[\.;,]$");

        /// <summary>
        /// Cấp bậc hàm: sĩ quan (Đại/Thượng/Trung/Thiếu + tướng/tá/úy), hạ sĩ quan (Thượng sĩ, Trung sĩ, Hạ sĩ),
        /// chiến sĩ (Binh nhất, Binh nhì).
        /// </summary>
        public static readonly Regex RxRank = new Regex(
            @"^((Đại|Thượng|Trung|Thiếu)\s+(tướng|tá|úy|uý)|(Thượng|Trung|Hạ)\s+sĩ|Binh\s+(nhất|nhì))\s+");

        /// <summary>
        /// Chức danh khoa học, học vị, danh hiệu ghi trước cấp bậc hàm (VBHN 08 Điều 13 khoản 1 điểm d):
        /// "PGS.TS Thiếu tướng Nguyễn Văn A", "GS, TS Đại tá ...", "NGND Đại tá ...".
        /// </summary>
        public static readonly Regex RxTitlePrefix = new Regex(
            @"^((GS|PGS|TSKH|TS|ThS|BSCKII|BSCKI|BS|DS|KS|NGND|NGƯT|NSND|NSƯT|AHLLVTND|AHLĐ)\.?\s*[,\.]?\s*)+(?=\S)");

        static string StripTitle(string s) => RxTitlePrefix.Replace(s, "");

        public static bool HasRank(string s) => RxRank.IsMatch(StripTitle(TextUtil.Clean(s)));

        public static bool IsPersonName(string s)
        {
            s = StripTitle(TextUtil.Clean(s));
            if (s.Length == 0 || s.Length > 50 || Regex.IsMatch(s, @"[\d:;,\.\(\)/]")) return false;
            // cấp bậc hàm đứng trước họ tên (Quân đội TT 68, Công an 08/VBHN-BCA)
            s = RxRank.Replace(s, "");
            var words = s.Split(' ');
            if (words.Length < 2 || words.Length > 7) return false;
            return words.All(w => w.Length > 0 && char.IsUpper(w[0]));
        }
    }
}
