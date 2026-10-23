"""Tạo samples/compare.html: ảnh trang đầu Trước | Sau + báo cáo của từng mẫu."""
import os, html, re, base64
root = os.path.join(os.path.dirname(__file__), "..", "samples")
def img(path):
    return "data:image/png;base64," + base64.b64encode(open(os.path.join(root, path), "rb").read()).decode()

names = sorted(f[:-5] for f in os.listdir(os.path.join(root, "input")) if f.endswith(".docx"))
parts = ["""<!doctype html><meta charset="utf-8"><title>V-Format — đầu vào / đầu ra</title>
<style>body{font:14px system-ui;margin:24px;background:#f4f6f9;color:#222}h2{margin-top:40px;color:#005a9e}
.row{display:flex;gap:16px}.col{flex:1;background:#fff;padding:8px;border:1px solid #d6dde6}.col img{width:100%;display:block}
.col h3{margin:4px 0 8px;font-size:13px;color:#555}details{background:#fff;border:1px solid #d6dde6;padding:8px 12px;margin-top:8px}
pre{white-space:pre-wrap;font:12px ui-monospace,monospace}.sum{background:#e8f1fb;padding:8px 12px;border-left:4px solid #005a9e}</style>
<h1>V-Format — kết quả test đầu vào / đầu ra</h1>
<p>Ảnh render bằng Quick Look của macOS (xấp xỉ Word; không vẽ viền bảng theo style, không ẩn số trang trang đầu, đặt sai vị trí cột phải của bảng chữ ký). Mở file .docx bằng Word để xem chính xác.</p>"""]
for n in names:
    rep = open(os.path.join(root, "output", n + ".md"), encoding="utf-8").read()
    summary = "<br>".join(html.escape(l.replace("**", "").strip()) for l in rep.splitlines() if l.startswith("**"))
    parts.append(f"""<h2>{n}</h2><div class="sum">{summary}</div><div class="row">
<div class="col"><h3>ĐẦU VÀO — input/{n}.docx</h3><img src="{img(f'render/input-{n}.png')}"></div>
<div class="col"><h3>ĐẦU RA — output/{n}.docx</h3><img src="{img(f'render/output-{n}.png')}"></div></div>
<details><summary>Báo cáo chi tiết (nhận diện từng đoạn, lỗi trước/sau)</summary><pre>{html.escape(rep)}</pre></details>""")
open(os.path.join(root, "compare.html"), "w", encoding="utf-8").write("\n".join(parts))
print("ok")
