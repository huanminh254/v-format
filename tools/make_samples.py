"""Tạo văn bản mẫu SAI thể thức để test V-Format (python3 tools/make_samples.py)."""
import os
from docx import Document
from docx.shared import Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH

OUT = os.path.join(os.path.dirname(__file__), "..", "samples", "input")
os.makedirs(OUT, exist_ok=True)

def new(font="Arial", size=12):
    d = Document()  # mặc định khổ Letter, lề 1 inch
    st = d.styles["Normal"]
    st.font.name = font
    st.font.size = Pt(size)
    return d

def add(d, text, bold=False, italic=False, size=None, align=None, font=None):
    p = d.add_paragraph()
    r = p.add_run(text)
    r.bold = bold; r.italic = italic
    if size: r.font.size = Pt(size)
    if font: r.font.name = font
    if align == "c": p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    if align == "r": p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    return p

LOREM = ("Các cơ quan, đơn vị có trách nhiệm tổ chức triển khai thực hiện nghiêm túc  nội dung này , bảo đảm đúng tiến độ, "
         "chất lượng và hiệu quả; định kỳ báo cáo kết quả thực hiện về cơ quan thường trực để tổng hợp, báo cáo cấp trên theo quy định.")

# 1. Quyết định hành chính: toàn đoạn rời, Arial 12, nhiều lỗi chữ
d = new()
for t in ["ỦY BAN NHÂN DÂN TỈNH HÀ NAM", "SỞ NỘI VỤ", "Số :12 / QĐ- SNV",
          "CỘNG HOÀ XÃ HỘI CHỦ NGHĨA VIỆT NAM", "Độc lập-Tự do-Hạnh phúc", "Hà Nam , ngày 5 tháng 1 năm 2026.", ""]:
    add(d, t, align="c" if t.isupper() else None)
add(d, "QUYẾT ĐỊNH", bold=True, align="c")
add(d, "Về việc ban hành Quy chế làm việc của Sở Nội vụ", align="c")
add(d, "GIÁM ĐỐC SỞ NỘI VỤ TỈNH HÀ NAM", bold=True, align="c")
add(d, "Căn cứ Luật Tổ chức chính quyền địa phương ngày 19 tháng 6 năm 2015;")
add(d, "Căn cứ Quyết định số 25/2021/QĐ-UBND của uỷ ban nhân dân tỉnh về việc sử lý hồ sơ;")
add(d, "Theo đề nghị của Chánh Văn phòng Sở.")
add(d, "QUYẾT ĐỊNH:", bold=True, align="c")
add(d, "Điều 1 : Ban hành kèm theo Quyết định này Quy chế làm việc của Sở Nội vụ tỉnh Hà Nam.")
add(d, "Điều 2. Quyết định này có hiệu lực thi hành kể từ ngày ký.  " + LOREM)
add(d, "Điều 3. Chánh Văn phòng, Trưởng các phòng, đơn vị thuộc Sở và các cá nhân có liên quan chịu trách nhiệm thi hành Quyết định này./.")
add(d, "")
for t in ["Nơi nhận:", "+ Như Điều 3", "- UBND tỉnh (để b/c)", "- Lưu: VT, VP"]:
    add(d, t)
for t in ["KT. GIÁM ĐỐC", "PHÓ GIÁM ĐỐC", "", "", "Trần Thị Bình"]:
    add(d, t, bold=t.isupper(), align="r")
d.save(os.path.join(OUT, "01-quyet-dinh-hanh-chinh.docx"))

# 2. Công văn: khối đầu đã trong bảng có viền, cỡ chữ sai
d = new("Times New Roman", 13)
t = d.add_table(rows=1, cols=2); t.style = "Table Grid"
c1, c2 = t.rows[0].cells
c1.paragraphs[0].add_run("UBND HUYỆN KIM BẢNG")
c1.add_paragraph().add_run("PHÒNG GIÁO DỤC VÀ ĐÀO TẠO").bold = True
c1.add_paragraph().add_run("Số: 215/PGDĐT-VP")
c1.add_paragraph().add_run("V/v: tổ chức hội thi giáo viên dạy giỏi")
r = c2.paragraphs[0].add_run("CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM"); r.font.size = Pt(14)
c2.add_paragraph().add_run("Độc lập – Tự do – Hạnh phúc")
c2.add_paragraph().add_run("Kim Bảng, ngày 8 tháng 2 năm 2026").italic = False
add(d, "")
add(d, "Kính gửi : Hiệu trưởng các trường Mầm non, Tiểu học, THCS", align="c")
for i in range(3):
    add(d, LOREM)
add(d, "Trân trọng./.")
t2 = d.add_table(rows=1, cols=2)
a, b = t2.rows[0].cells
a.paragraphs[0].add_run("Nơi nhận:")
for x in ["- Như trên", "- Lưu: VT"]:
    a.add_paragraph().add_run(x)
b.paragraphs[0].add_run("TRƯỞNG PHÒNG")
b.add_paragraph(); b.add_paragraph()
b.add_paragraph().add_run("Nguyễn Văn Hùng")
d.save(os.path.join(OUT, "02-cong-van-co-bang.docx"))

# 3. Nghị quyết Đảng: đoạn rời, thiếu dấu *, "TM." thay vì "T/M"
d = new("Arial", 13)
for t in ["ĐẢNG BỘ TỈNH HÀ NAM", "HUYỆN ỦY KIM BẢNG", "Số: 15 - NQ/HU", "ĐẢNG CỘNG SẢN VIỆT NAM",
          "Kim Bảng, ngày 3 tháng 2 năm 2026", "", "NGHỊ QUYẾT", "về phát triển kinh tế - xã hội năm 2026", ""]:
    add(d, t, align="c")
for t in ["I. TÌNH HÌNH", LOREM, "II. MỤC TIÊU, NHIỆM VỤ", "1. Mục tiêu chung", LOREM, "2. Nhiệm vụ trọng tâm", LOREM,
          "III. TỔ CHỨC THỰC HIỆN", LOREM, ""]:
    add(d, t)
for t in ["Nơi nhận:", "- Tỉnh ủy (báo cáo)", "- Lưu Văn phòng Huyện ủy", "TM. Ban Chấp hành", "BÍ THƯ", "", "Nguyễn Văn An"]:
    add(d, t)
d.save(os.path.join(OUT, "03-nghi-quyet-dang.docx"))

# 4. Báo cáo Đoàn nhiều trang + bảng số liệu tràn lề
d = new("Arial", 12)
for t in ["BCH ĐOÀN TỈNH HÀ NAM", "BCH ĐOÀN HUYỆN KIM BẢNG", "***", "Số: 45-BC/HĐTN",
          "ĐOÀN TNCS HỒ CHÍ MINH", "Kim Bảng, ngày 20 tháng 12 năm 2026", "BÁO CÁO",
          "Kết quả công tác Đoàn và phong trào thanh thiếu nhi năm 2026", ""]:
    add(d, t, align="c")
add(d, "I. KẾT QUẢ ĐẠT ĐƯỢC")
add(d, "1. Công tác giáo dục chính trị tư tưởng")
for i in range(4): add(d, LOREM)
add(d, "2. Kết quả thực hiện các chỉ tiêu")
tbl = d.add_table(rows=6, cols=5); tbl.style = "Table Grid"
tbl.autofit = False
rows = [["STT", "Chỉ tiêu", "Kế hoạch", "Thực hiện", "Tỷ lệ (%)"], ["1", "Kết nạp đoàn viên mới", "500", "560", "112"],
        ["2", "Giới thiệu đoàn viên ưu tú cho Đảng", "120", "131", "109"], ["3", "Công trình thanh niên", "15", "17", "113"],
        ["4", "Thanh niên tình nguyện", "2.000", "2.350", "117"], ["5", "Ngày công lao động", "10.000", "11.200", "112"]]
for r_, row in enumerate(rows):
    for c_, v in enumerate(row):
        tbl.cell(r_, c_).text = v
        tbl.cell(r_, c_).width = Cm(5)  # 5 x 5 cm = 25 cm: tràn lề
for i in range(5): add(d, LOREM)
add(d, "II. TỒN TẠI, HẠN CHẾ")
for i in range(4): add(d, LOREM)
add(d, "III. PHƯƠNG HƯỚNG NĂM 2027")
for i in range(5): add(d, LOREM)
add(d, "")
for t in ["Nơi nhận:", "- Tỉnh Đoàn (b/c)", "- Lưu VP", "TM. BAN THƯỜNG VỤ HUYỆN ĐOÀN", "BÍ THƯ", "", "", "Lê Thị Mai"]:
    add(d, t)
d.save(os.path.join(OUT, "04-bao-cao-doan.docx"))

# 5. Kế hoạch bài dạy
d = new("Calibri", 12)
for t in ["Trường: THCS Lê Hồng Phong", "Tổ: Khoa học tự nhiên", "Họ và tên giáo viên: Phạm Thu Hà",
          "KẾ HOẠCH BÀI DẠY", "TÊN BÀI DẠY: TỐC ĐỘ CHUYỂN ĐỘNG", "Môn học: Khoa học tự nhiên; lớp: 7",
          "Thời gian thực hiện: 2 tiết", "I. MỤC TIÊU", "1. Về kiến thức", "Nêu được ý nghĩa vật lí của tốc độ.",
          "2. Về năng lực", "Năng lực tự chủ và tự học, giao tiếp và hợp tác.", "II. THIẾT BỊ DẠY HỌC VÀ HỌC LIỆU",
          "Máy chiếu, phiếu học tập.", "III. TIẾN TRÌNH DẠY HỌC", "1. Hoạt động 1: Mở đầu", "a) Mục tiêu",
          "Tạo hứng thú cho học sinh.", "b) Nội dung", "Học sinh quan sát video.", "c) Sản phẩm", "Câu trả lời của học sinh.",
          "d) Tổ chức thực hiện", "Giáo viên chiếu video, học sinh thảo luận nhóm."]:
    add(d, t)
d.save(os.path.join(OUT, "05-ke-hoach-bai-day.docx"))
# 6. Công văn Công an (08/VBHN-BCA): đoạn rời, sai chữ, người ký thiếu cấp bậc hàm
d = new("Arial", 13)
for t in ["CÔNG AN TỈNH HÀ NAM", "PHÒNG THAM MƯU", "Số:  215 /PV01-TH", "V/v : triển khai công tác bảo đảm an ninh trật tự dịp Tết",
          "CỘNG HOÀ XÃ HỘI CHỦ NGHĨA VIỆT NAM", "Độc lập - Tự do -  Hạnh phúc", "Hà Nam, ngày 9 tháng 2 năm 2026", ""]:
    add(d, t, align="c" if t.isupper() else None)
add(d, "Kính gửi : Công an các huyện, thị xã, thành phố")
for i in range(3): add(d, LOREM)
add(d, "Đề nghị các đơn vị nghiêm túc triển khai thực hiện./.")
add(d, "")
for t in ["Nơi nhận:", "- Như trên", "- Đ/c Giám đốc (để b/c)", "- Lưu: VT, PV01"]:
    add(d, t)
for t in ["KT. TRƯỞNG PHÒNG", "PHÓ TRƯỞNG PHÒNG", "", "", "Lê Văn Nam"]:
    add(d, t, bold=t.isupper(), align="r")
d.save(os.path.join(OUT, "06-cong-van-cong-an.docx"))
print("ok", sorted(os.listdir(OUT)))
