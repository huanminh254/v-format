using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;
using VFormat.Core;
using VFormat.Docx;

namespace VFormat.App
{
    /// <summary>Cửa sổ chính: kéo thả .docx → tạo bản chuẩn hóa; tạo văn bản mới từ thư viện mẫu.</summary>
    public class MainForm : Form
    {
        static readonly Color Accent = Color.FromArgb(0, 90, 158);
        static readonly Color Soft = Color.FromArgb(243, 246, 250);

        ComboBox _standard, _size, _spacing;
        RadioButton _sameFolder, _otherFolder;
        TextBox _folder, _diaDanh;
        Panel _drop;
        ListView _files;
        Label _status;
        ListBox _templates;
        readonly Queue<string> _queue = new Queue<string>();
        bool _busy;

        public MainForm()
        {
            Text = "V-FORMAT — Chuẩn hóa thể thức văn bản Word";
            Font = new Font("Segoe UI", 9.5f);
            MinimumSize = new Size(820, 620);
            Size = new Size(980, 720);
            StartPosition = FormStartPosition.CenterScreen;
            BackColor = Color.White;
            AllowDrop = true;
            DragEnter += OnDragEnter;
            DragDrop += OnDragDrop;
            BuildUi();
        }

        // ================= Giao diện =================

        void BuildUi()
        {
            var header = new Panel { Dock = DockStyle.Top, Height = 64, BackColor = Accent };
            header.Controls.Add(new Label
            {
                Text = "V-FORMAT", ForeColor = Color.White, Font = new Font("Segoe UI Semibold", 18f),
                AutoSize = true, Location = new Point(18, 6),
            });
            header.Controls.Add(new Label
            {
                Text = "Chuẩn hóa thể thức văn bản Word theo NĐ 30 · TT 68 · Công an · HD 05 · HD 29 · NĐ 250/2026 · CV 5512 — offline",
                ForeColor = Color.FromArgb(210, 228, 245), AutoSize = true, Location = new Point(21, 40),
            });

            var tabs = new TabControl { Dock = DockStyle.Fill, Padding = new Point(14, 6) };
            tabs.TabPages.Add(BuildFormatTab());
            tabs.TabPages.Add(BuildTemplateTab());

            _status = new Label
            {
                Dock = DockStyle.Bottom, Height = 26, TextAlign = ContentAlignment.MiddleLeft, Padding = new Padding(12, 0, 0, 0),
                BackColor = Soft, ForeColor = Color.DimGray, Text = "Sẵn sàng.",
            };
            Controls.Add(tabs);
            Controls.Add(_status);
            Controls.Add(header);
        }

        TabPage BuildFormatTab()
        {
            var page = new TabPage("  Chuẩn hóa văn bản  ") { BackColor = Color.White, Padding = new Padding(12) };

            // --- tùy chọn
            var opts = new TableLayoutPanel { Dock = DockStyle.Top, AutoSize = true, ColumnCount = 6, Padding = new Padding(0, 0, 0, 6) };
            _standard = Combo(260, new[] { "Tự nhận diện" }.Concat(Profiles.All.Select(StandardLabel)).ToArray());
            _size = Combo(70, "14", "13");
            _spacing = Combo(70, "1.15", "1.0", "1.3", "1.5");
            opts.Controls.Add(Caption("Quy chuẩn"));
            opts.Controls.Add(_standard);
            opts.Controls.Add(Caption("Cỡ chữ"));
            opts.Controls.Add(_size);
            opts.Controls.Add(Caption("Giãn dòng"));
            opts.Controls.Add(_spacing);

            var outRow = new FlowLayoutPanel { Dock = DockStyle.Top, AutoSize = true, WrapContents = false, Padding = new Padding(0, 0, 0, 8) };
            _sameFolder = new RadioButton { Text = "Lưu cạnh tệp gốc (thêm \"_chuan\")", AutoSize = true, Checked = true, Margin = new Padding(0, 6, 16, 0) };
            _otherFolder = new RadioButton { Text = "Lưu vào thư mục:", AutoSize = true, Margin = new Padding(0, 6, 4, 0) };
            _folder = new TextBox { Width = 300, Enabled = false, Text = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) };
            var browse = SmallButton("…");
            browse.Enabled = false;
            _otherFolder.CheckedChanged += (s, e) => { _folder.Enabled = browse.Enabled = _otherFolder.Checked; };
            browse.Click += (s, e) =>
            {
                using (var d = new FolderBrowserDialog { SelectedPath = _folder.Text })
                    if (d.ShowDialog(this) == DialogResult.OK) _folder.Text = d.SelectedPath;
            };
            outRow.Controls.AddRange(new Control[] { _sameFolder, _otherFolder, _folder, browse });

            // --- vùng thả
            _drop = new Panel { Dock = DockStyle.Top, Height = 150, Cursor = Cursors.Hand, AllowDrop = true, Margin = new Padding(0, 6, 0, 6) };
            _drop.Paint += PaintDropZone;
            _drop.Click += (s, e) => PickFiles();
            _drop.DragEnter += OnDragEnter;
            _drop.DragDrop += OnDragDrop;
            _drop.DragLeave += (s, e) => { _drop.Tag = null; _drop.Invalidate(); };
            _drop.Resize += (s, e) => _drop.Invalidate();

            // --- danh sách kết quả
            _files = new ListView
            {
                Dock = DockStyle.Fill, View = View.Details, FullRowSelect = true, GridLines = false, HideSelection = false,
            };
            _files.Columns.Add("Tệp", 300);
            _files.Columns.Add("Quy chuẩn", 170);
            _files.Columns.Add("Lỗi thể thức", 110);
            _files.Columns.Add("Trạng thái", 330);
            _files.DoubleClick += (s, e) => OpenSelected();

            var actions = new FlowLayoutPanel { Dock = DockStyle.Bottom, AutoSize = true, Padding = new Padding(0, 6, 0, 0) };
            var open = ActionButton("Mở tệp đã chuẩn hóa", primary: true); open.Click += (s, e) => OpenSelected();
            var folder = ActionButton("Mở thư mục"); folder.Click += (s, e) => OpenFolder();
            var report = ActionButton("Xem báo cáo"); report.Click += (s, e) => ShowReport();
            var clear = ActionButton("Xóa danh sách"); clear.Click += (s, e) => { if (!_busy) _files.Items.Clear(); };
            actions.Controls.AddRange(new Control[] { open, folder, report, clear });

            var spacer = new Panel { Dock = DockStyle.Top, Height = 10 };
            page.Controls.Add(_files);
            page.Controls.Add(spacer);
            page.Controls.Add(_drop);
            page.Controls.Add(outRow);
            page.Controls.Add(opts);
            page.Controls.Add(actions);
            return page;
        }

        TabPage BuildTemplateTab()
        {
            var page = new TabPage("  Tạo văn bản từ mẫu  ") { BackColor = Color.White, Padding = new Padding(12) };
            var hint = new Label
            {
                Dock = DockStyle.Top, Height = 34, ForeColor = Color.DimGray,
                Text = $"Chọn một trong {TemplateLibrary.All.Count} mẫu dựng sẵn — V-Format tạo tệp .docx đã đúng thể thức để bạn điền nội dung.",
            };
            _templates = new ListBox { Dock = DockStyle.Fill, IntegralHeight = false, Font = new Font("Segoe UI", 10f) };
            foreach (var t in TemplateLibrary.All) _templates.Items.Add(t);
            _templates.SelectedIndex = 0;
            _templates.DoubleClick += (s, e) => CreateFromTemplate();

            var bottom = new FlowLayoutPanel { Dock = DockStyle.Bottom, AutoSize = true, Padding = new Padding(0, 8, 0, 0) };
            _diaDanh = new TextBox { Width = 180, Text = "Hà Nội" };
            var create = ActionButton("Tạo văn bản…", primary: true);
            create.Click += (s, e) => CreateFromTemplate();
            bottom.Controls.Add(new Label { Text = "Địa danh:", AutoSize = true, Margin = new Padding(0, 8, 4, 0) });
            bottom.Controls.Add(_diaDanh);
            bottom.Controls.Add(create);

            page.Controls.Add(_templates);
            page.Controls.Add(hint);
            page.Controls.Add(bottom);
            return page;
        }

        void PaintDropZone(object sender, PaintEventArgs e)
        {
            var g = e.Graphics;
            g.SmoothingMode = SmoothingMode.AntiAlias;
            var r = new Rectangle(1, 1, _drop.Width - 3, _drop.Height - 3);
            bool hot = _drop.Tag != null;
            using (var bg = new SolidBrush(hot ? Color.FromArgb(225, 238, 252) : Soft)) g.FillRectangle(bg, r);
            using (var pen = new Pen(hot ? Accent : Color.FromArgb(160, 180, 205), 2) { DashStyle = DashStyle.Dash }) g.DrawRectangle(pen, r);
            var big = "Kéo thả tệp Word (.docx) hoặc thư mục vào đây";
            var small = _busy ? "Đang xử lý…" : "hoặc bấm để chọn tệp — có thể chọn nhiều tệp một lúc";
            using (var f1 = new Font("Segoe UI Semibold", 14f))
            using (var b1 = new SolidBrush(Accent))
            using (var b2 = new SolidBrush(Color.DimGray))
            {
                var s1 = g.MeasureString(big, f1); var s2 = g.MeasureString(small, Font);
                g.DrawString(big, f1, b1, (r.Width - s1.Width) / 2, r.Height / 2f - s1.Height);
                g.DrawString(small, Font, b2, (r.Width - s2.Width) / 2, r.Height / 2f + 4);
            }
        }

        static Label Caption(string t) => new Label { Text = t, AutoSize = true, Margin = new Padding(0, 7, 6, 0) };

        static ComboBox Combo(int width, params string[] items)
        {
            var c = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList, Width = width, Margin = new Padding(0, 3, 18, 3) };
            c.Items.AddRange(items);
            c.SelectedIndex = 0;
            return c;
        }

        static Button SmallButton(string t) => new Button { Text = t, Width = 32, Height = 26, FlatStyle = FlatStyle.System };

        static Button ActionButton(string text, bool primary = false)
        {
            var b = new Button
            {
                Text = text, AutoSize = true, Height = 34, Padding = new Padding(10, 0, 10, 0), FlatStyle = FlatStyle.Flat,
                Margin = new Padding(0, 0, 8, 0), Cursor = Cursors.Hand,
                BackColor = primary ? Accent : Soft, ForeColor = primary ? Color.White : Color.Black,
            };
            b.FlatAppearance.BorderColor = primary ? Accent : Color.FromArgb(205, 214, 226);
            return b;
        }

        static string StandardLabel(StandardKind k)
        {
            var p = Profiles.Get(k);
            return $"{p.ShortName} — {p.Description}";
        }

        FormatOptions Options() => new FormatOptions
        {
            Standard = _standard.SelectedIndex <= 0 ? (StandardKind?)null : Profiles.All[_standard.SelectedIndex - 1],
            BodySize = float.Parse((string)_size.SelectedItem, System.Globalization.CultureInfo.InvariantCulture),
            LineSpacing = float.Parse((string)_spacing.SelectedItem, System.Globalization.CultureInfo.InvariantCulture),
        };

        // ================= Kéo thả / chọn tệp =================

        void OnDragEnter(object sender, DragEventArgs e)
        {
            if (!e.Data.GetDataPresent(DataFormats.FileDrop)) return;
            e.Effect = DragDropEffects.Copy;
            _drop.Tag = true; _drop.Invalidate();
        }

        void OnDragDrop(object sender, DragEventArgs e)
        {
            _drop.Tag = null; _drop.Invalidate();
            if (e.Data.GetData(DataFormats.FileDrop) is string[] paths) Enqueue(paths);
        }

        void PickFiles()
        {
            if (_busy) return;
            using (var d = new OpenFileDialog { Filter = "Văn bản Word (*.docx)|*.docx|Mọi tệp|*.*", Multiselect = true, Title = "Chọn văn bản cần chuẩn hóa" })
                if (d.ShowDialog(this) == DialogResult.OK) Enqueue(d.FileNames);
        }

        public void Enqueue(IEnumerable<string> paths)
        {
            foreach (var p in paths)
            {
                if (Directory.Exists(p))
                {
                    foreach (var f in Directory.GetFiles(p, "*.doc*", SearchOption.TopDirectoryOnly)) EnqueueFile(f);
                }
                else if (File.Exists(p)) EnqueueFile(p);
            }
            _ = ProcessQueue();
        }

        void EnqueueFile(string f)
        {
            var name = Path.GetFileName(f);
            if (name.StartsWith("~$")) return;                                         // tệp khóa tạm của Word
            if (Path.GetFileNameWithoutExtension(f).EndsWith("_chuan", StringComparison.OrdinalIgnoreCase)) return; // đã là đầu ra
            _queue.Enqueue(f);
            var item = new ListViewItem(new[] { name, "", "", "Đang chờ…" }) { Tag = new Row { Input = f } };
            _files.Items.Add(item);
        }

        class Row { public string Input; public FormatResult Result; public string Error; }

        async Task ProcessQueue()
        {
            if (_busy) return;
            _busy = true; _drop.Invalidate();
            int ok = 0, fail = 0;
            try
            {
                while (_queue.Count > 0)
                {
                    var input = _queue.Dequeue();
                    var item = _files.Items.Cast<ListViewItem>().Last(i => ((Row)i.Tag).Input == input && ((Row)i.Tag).Result == null && ((Row)i.Tag).Error == null);
                    var row = (Row)item.Tag;
                    item.SubItems[3].Text = "Đang chuẩn hóa…";
                    _status.Text = "Đang xử lý " + Path.GetFileName(input);
                    var opts = Options();
                    var output = OutputPath(input);
                    try
                    {
                        row.Result = await Task.Run(() => DocxFormatter.Run(input, output, opts));
                        var r = row.Result;
                        item.SubItems[1].Text = r.Profile.ShortName + " — " + r.Profile.Description;
                        item.SubItems[2].Text = $"{r.IssuesBefore.Count} → {r.IssuesAfter.Count}";
                        item.SubItems[3].Text = "✓ " + Path.GetFileName(output);
                        item.ForeColor = r.IssuesAfter.Count == 0 && r.SchemaErrors == 0 ? Color.FromArgb(0, 120, 60) : Color.DarkOrange;
                        ok++;
                    }
                    catch (Exception ex)
                    {
                        row.Error = Friendly(ex);
                        item.SubItems[3].Text = "✗ " + row.Error;
                        item.ForeColor = Color.Firebrick;
                        fail++;
                    }
                }
            }
            finally
            {
                _busy = false; _drop.Invalidate();
                _status.Text = $"Xong: {ok} tệp đã chuẩn hóa" + (fail > 0 ? $", {fail} tệp lỗi." : ".") + "  Bấm đúp một dòng để mở tệp.";
            }
        }

        static string Friendly(Exception ex)
        {
            if (ex is NotSupportedException) return ex.Message;
            if (ex is IOException && ex.Message.Contains("being used")) return "Tệp đang mở trong Word — hãy đóng tệp rồi thử lại.";
            if (ex is UnauthorizedAccessException) return "Không có quyền ghi vào thư mục đích.";
            if (ex is DocumentFormat.OpenXml.Packaging.OpenXmlPackageException || ex is InvalidDataException)
                return "Tệp không phải .docx hợp lệ hoặc đã bị hỏng.";
            return ex.GetType().Name + ": " + ex.Message;
        }

        string OutputPath(string input)
        {
            var dir = _otherFolder.Checked && Directory.Exists(_folder.Text) ? _folder.Text : Path.GetDirectoryName(input);
            return Path.Combine(dir, Path.GetFileNameWithoutExtension(input) + "_chuan.docx");
        }

        // ================= Hành động =================

        Row Selected() => _files.SelectedItems.Count > 0 ? (Row)_files.SelectedItems[0].Tag : null;

        void OpenSelected()
        {
            var r = Selected();
            if (r?.Result == null) return;
            Process.Start(new ProcessStartInfo(r.Result.Output) { UseShellExecute = true });
        }

        void OpenFolder()
        {
            var r = Selected();
            var path = r?.Result?.Output ?? r?.Input;
            if (path == null) return;
            Process.Start("explorer.exe", "/select,\"" + path + "\"");
        }

        void ShowReport()
        {
            var r = Selected();
            if (r == null) return;
            var text = r.Result?.Report ?? r.Error ?? "";
            using (var f = new Form { Text = "Báo cáo — " + Path.GetFileName(r.Input), Size = new Size(900, 640), StartPosition = FormStartPosition.CenterParent })
            {
                f.Controls.Add(new TextBox
                {
                    Multiline = true, ReadOnly = true, ScrollBars = ScrollBars.Both, WordWrap = false, Dock = DockStyle.Fill,
                    Font = new Font("Consolas", 9.5f), Text = text.Replace("\n", "\r\n"), BackColor = Color.White,
                });
                f.ShowDialog(this);
            }
        }

        void CreateFromTemplate()
        {
            if (!(_templates.SelectedItem is DocTemplate t)) return;
            using (var d = new SaveFileDialog
            {
                Filter = "Văn bản Word (*.docx)|*.docx", FileName = t.Name.Replace("/", "-") + ".docx",
                InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments),
            })
            {
                if (d.ShowDialog(this) != DialogResult.OK) return;
                try
                {
                    Cursor = Cursors.WaitCursor;
                    var opts = Options();
                    var res = DocxFormatter.CreateFromTemplate(t, d.FileName, opts, string.IsNullOrWhiteSpace(_diaDanh.Text) ? null : _diaDanh.Text.Trim());
                    _status.Text = $"Đã tạo {Path.GetFileName(res.Output)} theo {res.Profile.FullName}.";
                    Process.Start(new ProcessStartInfo(res.Output) { UseShellExecute = true });
                }
                catch (Exception ex)
                {
                    MessageBox.Show(this, Friendly(ex), "V-FORMAT", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
                finally { Cursor = Cursors.Default; }
            }
        }
    }
}
