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
        static readonly Regex RxSoMuc = new Regex(@"^\d+(\.\d+)*\s*[\.\)]\s*\S");
        static readonly Regex RxHoatDong = new Regex(@"^Hoạt động\s+\d+", RegexOptions.IgnoreCase);
        static readonly Regex RxChuCai = new Regex(@"^[a-zđ]\s*[\)\.]\s+\S");
        static readonly Regex RxDecor = new Regex(@"^[\-_–—=\*\.]{2,}$|^\*$");
        static readonly Regex RxChucVuPrefix = new Regex(@"^(T\.?M\.?|T/M|K\.?T\.?|Q\.|T\.?L\.?|TUQ\.?|THỪA LỆNH|THỪA ỦY QUYỀN|THỪA UỶ QUYỀN)\s", RegexOptions.IgnoreCase);
        static readonly string[] ChucVuKeys =
        {
            "CHU TICH", "GIAM DOC", "BO TRUONG", "THU TRUONG", "BI THU", "TRUONG", "PHO ", "CHI HUY", "TU LENH",
            "CHINH UY", "HIEU TRUONG", "TONG ", "THU TUONG", "NGUOI ", "CHANH ", "VIEN TRUONG", "DAI DIEN",
            "UY VIEN", "THU KY", "CHU NHIEM", "CHU TOA", "GIAO VIEN", "XAC NHAN", "KE TOAN", "TAP THE",
        };

        public static ClassifiedDocument Classify(IList<ParagraphInfo> paragraphs, StandardKind standard)
        {
            var doc = new ClassifiedDocument { Standard = standard };
            doc.Paragraphs.AddRange(paragraphs);
            int n = paragraphs.Count;
            var t = paragraphs.Select(p => TextUtil.Clean(p.Text)).ToArray();
            var key = t.Select(TextUtil.Key).ToArray();
            var k = new ElementKind?[n];

            // 1. Mẫu độc lập ngữ cảnh
            for (int i = 0; i < n; i++)
            {
                if (t[i].Length == 0) { k[i] = ElementKind.Empty; continue; }
                var s = t[i]; var ky = key[i];
                if (ky == "CONG HOA XA HOI CHU NGHIA VIET NAM") k[i] = ElementKind.QuocHieu;
                else if (ky.StartsWith("DOC LAP") && ky.Contains("TU DO") && ky.Length < 45) k[i] = ElementKind.TieuNgu;
                else if (ky == "DANG CONG SAN VIET NAM") k[i] = ElementKind.TieuDeDang;
                else if (Regex.IsMatch(ky, @"^DOAN (TNCS|THANH NIEN CONG SAN) HO CHI MINH$")) k[i] = ElementKind.TieuDeDoan;
                else if (RxDecor.IsMatch(s)) k[i] = ElementKind.DauSao;
                else if (RxSo.IsMatch(s) && s.Length < 60) k[i] = ElementKind.SoKyHieu;
                else if (ky.StartsWith("V/V") || ky.StartsWith("VE VIEC ") && i < 15 && paragraphs[i].InTable) k[i] = ElementKind.TrichYeuCongVan;
                else if (RxNgayThang.IsMatch(s) && s.Length < 90 && s.IndexOf(',') > 0 && (i < 25 || paragraphs[i].InTable)) k[i] = ElementKind.DiaDanhNgayThang;
                else if (ky.StartsWith("KINH GUI")) k[i] = ElementKind.KinhGui;
                else if (Regex.IsMatch(ky, @"^NOI NHAN\s*:?$")) k[i] = ElementKind.NoiNhanTitle;
            }

            // 2. Vùng khối đầu
            var headerKinds = new[] { ElementKind.QuocHieu, ElementKind.TieuNgu, ElementKind.TieuDeDang, ElementKind.TieuDeDoan,
                                      ElementKind.SoKyHieu, ElementKind.DiaDanhNgayThang, ElementKind.TrichYeuCongVan };
            int headerEnd = -1;
            for (int i = 0; i < n && i < 20; i++)
                if (k[i].HasValue && headerKinds.Contains(k[i].Value)) headerEnd = i;
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
                    if (k[i].HasValue) continue;
                    if (TextUtil.IsAllUpper(t[i]) && paragraphs[i].CellColumn <= 1) coQuan.Add(i);
                    else if (i > 0 && k[i - 1] == ElementKind.TrichYeuCongVan) k[i] = ElementKind.TrichYeuCongVan;
                    else if (TextUtil.IsAllUpper(t[i])) coQuan.Add(i);
                }
                // Cơ quan đứng trước số ký hiệu / dấu sao
                for (int j = 0; j < coQuan.Count; j++)
                    k[coQuan[j]] = j == coQuan.Count - 1 ? ElementKind.CoQuanBanHanh : ElementKind.CoQuanChuQuan;
            }

            // 3. Tên loại + trích yếu
            int tenLoai = -1, seen = 0;
            for (int i = headerEnd + 1; i < n && seen < 6; i++)
            {
                if (k[i] == ElementKind.Empty || k[i] == ElementKind.DauSao) continue;
                seen++;
                if (k[i].HasValue) break;
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
                    if (count > 0 && s.EndsWith(".") && s.Length > 120) break;
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
            int lastChucVu = -1;
            for (int i = tailStart; i < n; i++)
            {
                if (k[i].HasValue) continue;
                var s = t[i];
                bool prefix = RxChucVuPrefix.IsMatch(s);   // "TM. Ban Chấp hành" viết thường vẫn là quyền hạn ký
                if (!(TextUtil.IsAllUpper(s) || prefix) || TextUtil.WordCount(s) > 14 || s.EndsWith(".") && !prefix) continue;
                var prev = PrevNonEmpty(k, i);
                if (prev.HasValue && (prev == ElementKind.Chuong || prev == ElementKind.Phan || prev == ElementKind.Muc)) continue;
                bool looks = RxChucVuPrefix.IsMatch(s) || ChucVuKeys.Any(c => key[i].Contains(c));
                bool afterBody = paragraphs[i].InTable || NextNonEmptyIsNameOrUpper(t, i) || i > n - 12;
                if (looks && afterBody)
                {
                    k[i] = ElementKind.ChucVuKy; lastChucVu = i;
                    if (paragraphs[i].InTable) signTables.Add(paragraphs[i].TableIndex);
                }
            }
            if (noiNhan >= 0 && paragraphs[noiNhan].InTable) signTables.Add(paragraphs[noiNhan].TableIndex);
            if (lastChucVu >= 0)
            {
                // Chức vụ kế tiếp ngay chức vụ (ví dụ "KT. GIÁM ĐỐC" / "PHÓ GIÁM ĐỐC") và họ tên
                for (int i = lastChucVu + 1; i < n; i++)
                {
                    if (k[i] == ElementKind.Empty) continue;
                    if (k[i].HasValue) break;
                    if (IsPersonName(t[i])) { k[i] = ElementKind.HoTenKy; break; }
                    if (Regex.IsMatch(t[i], @"^\(.*\)$")) continue; // (Ký, ghi rõ họ tên) / (Đã ký)
                    break;
                }
            }
            // Đoạn còn lại trong bảng chữ ký
            for (int i = 0; i < n; i++)
            {
                if (k[i].HasValue || !paragraphs[i].InTable || !signTables.Contains(paragraphs[i].TableIndex)) continue;
                if (paragraphs[i].CellColumn <= 1) k[i] = ElementKind.NoiNhanItem;
                else if (TextUtil.IsAllUpper(t[i])) k[i] = ElementKind.ChucVuKy;
                else if (IsPersonName(t[i])) k[i] = ElementKind.HoTenKy;
                else k[i] = ElementKind.TableText;
            }

            // 6. Cấu trúc nội dung
            bool afterStructural = false;
            for (int i = 0; i < n; i++)
            {
                if (k[i].HasValue)
                {
                    if (k[i] != ElementKind.Empty) afterStructural = false;
                    continue;
                }
                var s = t[i];
                ElementKind kind;
                if (RxPhan.IsMatch(s) && s.Length < 120) kind = ElementKind.Phan;
                else if (RxChuong.IsMatch(s) && s.Length < 160) kind = ElementKind.Chuong;
                else if (RxMuc.IsMatch(s) && s.Length < 160) kind = ElementKind.Muc;
                else if (afterStructural && TextUtil.IsAllUpper(s)) kind = ElementKind.TenPhanChuong;
                else if (paragraphs[i].InTable) kind = ElementKind.TableText;
                else if (RxDieu.IsMatch(s)) kind = ElementKind.Dieu;
                else if (key[i].StartsWith("CAN CU ") && i < bodyStart + 15) kind = ElementKind.CanCu;
                else if (RxLaMa.IsMatch(s) && s.Length < 200) kind = ElementKind.MucLaMa;
                else if ((RxSoMuc.IsMatch(s) || RxHoatDong.IsMatch(s)) && IsHeadingLike(s)) kind = ElementKind.MucSo;
                else if (RxChuCai.IsMatch(s) && IsHeadingLike(s) && s.Length < 90) kind = ElementKind.MucChuCai;
                else if (TextUtil.IsAllUpper(s) && TextUtil.WordCount(s) <= 12 && i >= bodyStart) kind = ElementKind.TenPhanChuong;
                else kind = ElementKind.Body;
                k[i] = kind;
                afterStructural = kind == ElementKind.Phan || kind == ElementKind.Chuong || kind == ElementKind.Muc;
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

        public static bool IsPersonName(string s)
        {
            s = TextUtil.Clean(s);
            if (s.Length == 0 || s.Length > 50 || Regex.IsMatch(s, @"[\d:;,\.\(\)/]")) return false;
            // cấp bậc quân hàm đứng trước họ tên (TT 68)
            s = Regex.Replace(s, @"^(Đại|Thượng|Trung|Thiếu)\s+(tướng|tá|úy|uý)\s+", "");
            var words = s.Split(' ');
            if (words.Length < 2 || words.Length > 7) return false;
            return words.All(w => w.Length > 0 && char.IsUpper(w[0]));
        }
    }
}
