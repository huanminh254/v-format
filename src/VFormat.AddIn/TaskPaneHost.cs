using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using VFormat.Core;

namespace VFormat.AddIn
{
    /// <summary>Bảng điều khiển V-Format (dock bên phải cửa sổ Word). Được Word tạo như ActiveX qua ProgId.</summary>
    [ComVisible(true)]
    [Guid(ClassId)]
    [ProgId(ProgIdName)]
    [ClassInterface(ClassInterfaceType.AutoDispatch)]
    public class TaskPaneHost : UserControl
    {
        public const string ClassId = "70138569-6BD8-4684-A0EE-C3F1BB5AAAFD";
        public const string ProgIdName = "VFormat.TaskPaneHost";

        static readonly Color Accent = Color.FromArgb(0, 90, 158);
        static readonly Font UiFont = new Font("Segoe UI", 9f);
        static readonly Font UiBold = new Font("Segoe UI Semibold", 9f);

        Connect _addin;
        ComboBox _standard, _size, _spacing;
        TextBox _diaDanh;
        CheckBox _fixText, _tables, _lines, _pageNo, _fmtTables;
        ListView _results;
        Label _status, _resultTitle;
        bool _loading;

        public TaskPaneHost()
        {
            Font = UiFont;
            BackColor = Color.White;
            AutoScroll = true;
            BuildUi();
        }

        internal void Initialize(Connect addin)
        {
            _addin = addin;
            LoadSettings();
            Settings.Changed += () => { if (!IsDisposed) BeginInvokeSafe(LoadSettings); };
        }

        void BeginInvokeSafe(Action a)
        {
            if (IsHandleCreated) BeginInvoke(a); else a();
        }

        void BuildUi()
        {
            var root = new TableLayoutPanel
            {
                Dock = DockStyle.Fill, ColumnCount = 1, Padding = new Padding(10), AutoSize = false,
            };
            root.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            Controls.Add(root);

            var header = new Label
            {
                Text = "V-FORMAT", Font = new Font("Segoe UI Semibold", 14f), ForeColor = Accent, AutoSize = true,
                Margin = new Padding(0, 0, 0, 0),
            };
            var sub = new Label
            {
                Text = "Chuẩn hóa thể thức văn bản 1-Click", ForeColor = Color.DimGray, AutoSize = true, Margin = new Padding(0, 0, 0, 8),
            };
            root.Controls.Add(header);
            root.Controls.Add(sub);

            root.Controls.Add(Section("Quy chuẩn"));
            _standard = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList, Dock = DockStyle.Top };
            _standard.Items.Add("Tự nhận diện");
            foreach (var k in Profiles.All) _standard.Items.Add(Connect.StandardLabel(k));
            _standard.SelectedIndexChanged += (s, e) => SaveSettings();
            root.Controls.Add(_standard);

            var row = new TableLayoutPanel { ColumnCount = 4, Dock = DockStyle.Top, AutoSize = true, Margin = new Padding(0, 6, 0, 0) };
            _size = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList, Width = 60 };
            _size.Items.AddRange(new object[] { "13", "14" });
            _spacing = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList, Width = 60 };
            _spacing.Items.AddRange(new object[] { "1.0", "1.15", "1.3", "1.5" });
            _size.SelectedIndexChanged += (s, e) => SaveSettings();
            _spacing.SelectedIndexChanged += (s, e) => SaveSettings();
            row.Controls.Add(new Label { Text = "Cỡ chữ", AutoSize = true, Anchor = AnchorStyles.Left });
            row.Controls.Add(_size);
            row.Controls.Add(new Label { Text = "Giãn dòng", AutoSize = true, Anchor = AnchorStyles.Left });
            row.Controls.Add(_spacing);
            root.Controls.Add(row);

            var dd = new TableLayoutPanel { ColumnCount = 2, Dock = DockStyle.Top, AutoSize = true };
            dd.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            dd.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            _diaDanh = new TextBox { Dock = DockStyle.Fill };
            _diaDanh.Leave += (s, e) => SaveSettings();
            dd.Controls.Add(new Label { Text = "Địa danh", AutoSize = true, Anchor = AnchorStyles.Left });
            dd.Controls.Add(_diaDanh);
            root.Controls.Add(dd);

            root.Controls.Add(Section("Tùy chọn chuẩn hóa"));
            _fixText = Check("Sửa chữ Quốc hiệu, số ký hiệu, ngày tháng, nơi nhận");
            _tables = Check("Tự dàn khối đầu, chữ ký thành bảng 2 cột");
            _lines = Check("Kẻ đường dưới tiêu ngữ, tên cơ quan, trích yếu");
            _pageNo = Check("Đánh số trang (ẩn trang đầu)");
            _fmtTables = Check("Bảng vừa khổ, lặp tiêu đề khi sang trang");
            foreach (var c in new[] { _fixText, _tables, _lines, _pageNo, _fmtTables }) root.Controls.Add(c);

            root.Controls.Add(Section("Thao tác"));
            var main = Button("⚡  Chuẩn hóa 1-Click", primary: true);
            main.Click += (s, e) => Run(() => _addin.RunFormat());
            root.Controls.Add(main);

            var grid = new TableLayoutPanel { ColumnCount = 2, Dock = DockStyle.Top, AutoSize = true };
            grid.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50));
            grid.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50));
            void G(string text, string id) { var b = Button(text); b.Click += (s, e) => Run(() => _addin.Execute(id)); grid.Controls.Add(b); }
            G("Rà soát thể thức", "btnReview");
            G("✕ Xóa nhận xét V-Format", "btnClearVF");
            G("Kiểm tra chính tả", "btnSpell");
            G("Sửa chính tả", "btnSpellFix");
            G("Chèn khối đầu", "btnInsHeader");
            G("Chèn chữ ký", "btnInsSign");
            G("Lề trang A4", "btnPageSetup");
            G("Đánh số trang", "btnPageNumber");
            root.Controls.Add(grid);

            _resultTitle = Section("Kết quả");
            root.Controls.Add(_resultTitle);
            _results = new ListView
            {
                View = View.Details, FullRowSelect = true, Dock = DockStyle.Top, Height = 260, HeaderStyle = ColumnHeaderStyle.Nonclickable,
                ShowItemToolTips = true,
            };
            _results.Columns.Add("Mã", 64);
            _results.Columns.Add("Thành phần", 90);
            _results.Columns.Add("Nội dung", 400);
            _results.ItemActivate += (s, e) => GoTo();
            _results.SelectedIndexChanged += (s, e) => GoTo();
            root.Controls.Add(_results);

            _status = new Label { AutoSize = true, ForeColor = Color.DimGray, MaximumSize = new Size(300, 0), Margin = new Padding(0, 6, 0, 0) };
            root.Controls.Add(_status);
        }

        static Label Section(string text) => new Label
        {
            Text = text.ToUpperInvariant(), Font = UiBold, ForeColor = Accent, AutoSize = true, Margin = new Padding(0, 12, 0, 4),
        };

        CheckBox Check(string text)
        {
            var c = new CheckBox { Text = text, AutoSize = true, Margin = new Padding(0, 1, 0, 1) };
            c.CheckedChanged += (s, e) => SaveSettings();
            return c;
        }

        static Button Button(string text, bool primary = false)
        {
            var b = new Button
            {
                Text = text, Dock = DockStyle.Fill, Height = primary ? 40 : 30, FlatStyle = FlatStyle.Flat,
                Margin = new Padding(0, 2, 4, 2), Cursor = Cursors.Hand,
            };
            if (primary)
            {
                b.BackColor = Accent; b.ForeColor = Color.White; b.Font = new Font("Segoe UI Semibold", 10f);
                b.FlatAppearance.BorderSize = 0;
            }
            else
            {
                b.BackColor = Color.FromArgb(243, 246, 250);
                b.FlatAppearance.BorderColor = Color.FromArgb(210, 218, 228);
            }
            return b;
        }

        void LoadSettings()
        {
            _loading = true;
            var s = Settings.Current;
            _standard.SelectedIndex = s.AutoDetect ? 0 : (int)s.Standard + 1;
            _size.SelectedItem = s.BodySize.ToString("0");
            _spacing.SelectedItem = s.LineSpacing.ToString("0.0#", System.Globalization.CultureInfo.InvariantCulture);
            if (_spacing.SelectedIndex < 0) _spacing.SelectedIndex = 1;
            _diaDanh.Text = s.DiaDanh;
            _fixText.Checked = s.FixText; _tables.Checked = s.BuildTables; _lines.Checked = s.DrawLines;
            _pageNo.Checked = s.PageNumbers; _fmtTables.Checked = s.FormatTables;
            _loading = false;
        }

        void SaveSettings()
        {
            if (_loading || _addin == null) return;
            var s = Settings.Current;
            s.AutoDetect = _standard.SelectedIndex <= 0;
            if (_standard.SelectedIndex > 0) s.Standard = Profiles.All[_standard.SelectedIndex - 1];
            if (_size.SelectedItem != null) s.BodySize = float.Parse((string)_size.SelectedItem);
            if (_spacing.SelectedItem != null)
                s.LineSpacing = float.Parse((string)_spacing.SelectedItem, System.Globalization.CultureInfo.InvariantCulture);
            s.DiaDanh = string.IsNullOrWhiteSpace(_diaDanh.Text) ? "Hà Nội" : _diaDanh.Text.Trim();
            s.FixText = _fixText.Checked; s.BuildTables = _tables.Checked; s.DrawLines = _lines.Checked;
            s.PageNumbers = _pageNo.Checked; s.FormatTables = _fmtTables.Checked;
            _loading = true;
            try { s.Save(); } finally { _loading = false; }
        }

        void Run(Action a)
        {
            if (_addin == null) return;
            SaveSettings();
            _addin.Safe(a);
        }

        internal void ShowResults(string title, List<ReviewItem> items)
        {
            _resultTitle.Text = $"{title.ToUpperInvariant()} ({items.Count})";
            _results.BeginUpdate();
            _results.Items.Clear();
            foreach (var it in items)
            {
                var lv = new ListViewItem(new[] { it.Code, it.Kind ?? "", it.Message }) { Tag = it, ToolTipText = it.Message };
                lv.ForeColor = it.Severity == Severity.Error ? Color.Firebrick : it.Severity == Severity.Info ? Color.DimGray : Color.FromArgb(160, 90, 0);
                _results.Items.Add(lv);
            }
            _results.EndUpdate();
            SetStatus(items.Count == 0 ? "Không phát hiện lỗi." : "Bấm vào một dòng để chuyển tới đoạn văn tương ứng.");
        }

        internal void SetStatus(string msg) => _status.Text = msg;

        void GoTo()
        {
            if (_addin == null || _results.SelectedItems.Count == 0) return;
            if (!(_results.SelectedItems[0].Tag is ReviewItem it) || it.Start < 0) return;
            _addin.Safe(() =>
            {
                var d = _addin.Doc();
                var r = d.Range(it.Start, Math.Max(it.Start, it.End - 1));
                r.Select();
                _addin.App.ActiveWindow.ScrollIntoView(r);
            });
        }
    }
}
