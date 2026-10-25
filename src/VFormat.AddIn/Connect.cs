using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Security;
using System.Text;
using System.Windows.Forms;
using Microsoft.Office.Core;
using VFormat.Core;
using Word = Microsoft.Office.Interop.Word;

namespace VFormat.AddIn
{
    /// <summary>
    /// COM add-in cho Microsoft Word (Office 2013+ / Microsoft 365, 32 và 64 bit).
    /// Đăng ký: installer/install.ps1 ghi CLSID vào HKCU và khóa Word\Addins\VFormat.Connect.
    /// </summary>
    [ComVisible(true)]
    [Guid(ClassId)]
    [ProgId(ProgIdName)]
    [ClassInterface(ClassInterfaceType.AutoDispatch)] // Ribbon gọi callback qua IDispatch
    public class Connect : IDTExtensibility2, IRibbonExtensibility, ICustomTaskPaneConsumer
    {
        public const string ClassId = "FA30749A-6CEA-4764-8B98-CA6CCD4F03C8";
        public const string ProgIdName = "VFormat.Connect";
        public const string Title = "V-FORMAT";

        Word.Application _app;
        IRibbonUI _ribbon;
        ICTPFactory _ctpFactory;
        readonly Dictionary<long, CustomTaskPane> _panes = new Dictionary<long, CustomTaskPane>();

        internal Word.Application App => _app;
        internal Settings Settings => Settings.Current;

        // ================= IDTExtensibility2 =================

        public void OnConnection(object application, ext_ConnectMode connectMode, object addInInst, ref Array custom)
        {
            _app = (Word.Application)application;
            // Script kiểm thử lấy add-in qua Application.COMAddIns("VFormat.Connect").Object
            try { ((COMAddIn)addInInst).Object = this; } catch { }
            try
            {
                ((Word.ApplicationEvents4_Event)_app).DocumentChange += () => _ribbon?.Invalidate();
                ((Word.ApplicationEvents4_Event)_app).WindowActivate += (d, w) => _ribbon?.Invalidate();
            }
            catch { }
            Settings.Changed += () => _ribbon?.Invalidate();
        }

        public void OnDisconnection(ext_DisconnectMode removeMode, ref Array custom)
        {
            foreach (var p in _panes.Values) try { p.Delete(); } catch { }
            _panes.Clear();
            _app = null;
            GC.Collect(); GC.WaitForPendingFinalizers();
        }

        public void OnAddInsUpdate(ref Array custom) { }
        public void OnStartupComplete(ref Array custom) { }
        public void OnBeginShutdown(ref Array custom) { }

        // ================= Ribbon =================

        public string GetCustomUI(string ribbonId)
        {
            using (var s = typeof(Connect).Assembly.GetManifestResourceStream("VFormat.AddIn.Ribbon.xml"))
            using (var r = new StreamReader(s, Encoding.UTF8))
                return r.ReadToEnd();
        }

        public void OnRibbonLoad(IRibbonUI ribbon) => _ribbon = ribbon;

        public void CTPFactoryAvailable(ICTPFactory factory) => _ctpFactory = factory;

        public void OnAction(IRibbonControl control) => Safe(() => Execute(control.Id));

        internal void Execute(string id)
        {
            switch (id)
            {
                case "btnFormat": RunFormat(); break;
                case "btnReview": RunReview(); break;
                case "btnClearVF": Info($"Đã xóa {Review.ClearComments(Doc(), true)} nhận xét của V-Format."); break;
                case "btnClearAll":
                    if (Confirm("Xóa TẤT CẢ nhận xét trong tài liệu (kể cả của người khác)?"))
                        Info($"Đã xóa {Review.ClearComments(Doc(), false)} nhận xét.");
                    break;
                case "btnPageSetup":
                {
                    var d = Doc();
                    var p = Insert.ProfileForActive(d, Settings);
                    using (new WordBatch(_app, "V-Format: Lề trang A4")) Formatter.ApplyPageSetup(d, p);
                    Status($"Khổ A4, lề trên {p.MarginTopMm} – dưới {p.MarginBottomMm} – trái {p.MarginLeftMm} – phải {p.MarginRightMm} mm ({p.ShortName}).");
                    break;
                }
                case "btnPageNumber":
                {
                    var d = Doc();
                    var p = Insert.ProfileForActive(d, Settings);
                    using (new WordBatch(_app, "V-Format: Đánh số trang")) PageNumbers.Apply(d, p);
                    Status("Đã đánh số trang (ẩn ở trang đầu).");
                    break;
                }
                case "btnTables": RunFormatTablesOnly(); break;
                case "btnInsHeader": Insert.Header(_app, Doc(), Settings); break;
                case "btnInsSign": Insert.Signature(_app, Doc(), Settings); break;
                case "btnInsNoiNhan": Insert.NoiNhan(_app, Doc(), Settings); break;
                case "btnInsDate": Insert.Date(_app, Doc(), Settings); break;
                case "btnSpell": RunSpell(false); break;
                case "btnSpellFix":
                    if (Confirm("Tự động sửa mọi lỗi chính tả V-Format phát hiện được? (có thể Undo)")) RunSpell(true);
                    break;
                case "btnDict": OpenDictionary(); break;
                case "btnSaveTemplate": SaveTemplate(); break;
                case "btnTemplateFolder":
                    Directory.CreateDirectory(Settings.TemplateFolder);
                    Process.Start("explorer.exe", "\"" + Settings.TemplateFolder + "\"");
                    break;
                case "btnManualHeader":
                    Formatter.FormatHeaderOnly(_app, Doc(), Settings);
                    Status("Đã chuẩn hóa riêng khối Tiêu đề (Quốc hiệu / Cơ quan).");
                    break;
                case "btnManualSign":
                    Formatter.FormatSignOnly(_app, Doc(), Settings);
                    Status("Đã chuẩn hóa riêng khối Chữ ký & Nơi nhận.");
                    break;
                case "btnManualBody":
                    Formatter.FormatBodyOnly(_app, Doc(), Settings);
                    Status("Đã chuẩn hóa riêng phần Thân bài.");
                    break;
                case "btnDeleteBlankPages":
                {
                    int del = WordToolHelper.DeleteBlankPages(Doc());
                    Info(del > 0 ? $"Đã xóa {del} dòng/trang trắng thừa cuối văn bản!" : "Không tìm thấy trang trắng thừa nào.");
                    break;
                }
                case "btnCleanSpaces":
                    WordToolHelper.CleanSpacingErrors(Doc(), _app.Selection);
                    Status("Đã làm sạch khoảng trắng lỗi và dấu cách trước dấu câu.");
                    break;
                case "btnNumberToWords":
                {
                    var sel = _app.Selection;
                    if (sel == null || sel.Start == sel.End)
                    {
                        Info("Vui lòng bôi đen số tiền cần chuyển sang chữ (ví dụ: 15.000.000) rồi bấm lại.");
                        break;
                    }
                    string selText = sel.Text ?? "";
                    if (VndAmountReader.TryConvert(selText, out var words, out var err))
                    {
                        sel.Text = words;
                        Status($"Đã chuyển số tiền: {words}");
                    }
                    else
                    {
                        Info($"Không thể chuyển số tiền: {err}");
                    }
                    break;
                }
                case "btnConvertOldFonts":
                    WordToolHelper.ConvertTcvn3ToUnicode(Doc());
                    Info("Đã chuyển đổi mã TCVN3 sang Unicode dựng sẵn thành công!");
                    break;
                case "insTM":
                    InsertTextAtSelection("TM. ỦY BAN NHÂN DÂN\nCHỦ TỊCH\n\n\n\n(Họ và tên)", true);
                    break;
                case "insKT":
                    InsertTextAtSelection("KT. CHỦ TỊCH\nPHÓ CHỦ TỊCH\n\n\n\n(Họ và tên)", true);
                    break;
                case "insTL":
                    InsertTextAtSelection("TL. CHỦ TỊCH\nCHÁNH VĂN PHÒNG\n\n\n\n(Họ và tên)", true);
                    break;
                case "insTUQ":
                    InsertTextAtSelection("TUQ. GIÁM ĐỐC\nTRƯỞNG PHÒNG\n\n\n\n(Họ và tên)", true);
                    break;
                case "insQ":
                    InsertTextAtSelection("QUYỀN GIÁM ĐỐC\n\n\n\n(Họ và tên)", true);
                    break;
                case "insHoaToc":
                    InsertStampAtSelection("HỎA TỐC");
                    break;
                case "insThuongKhan":
                    InsertStampAtSelection("THƯỢNG KHẨN");
                    break;
                case "insMat":
                    InsertStampAtSelection("MẬT");
                    break;
                case "insToiMat":
                    InsertStampAtSelection("TỐI MẬT");
                    break;
                case "btnAbout":
                    Info("V-FORMAT — Chuẩn hóa thể thức văn bản Word 1-Click\n\n" +
                         string.Join("\n", Profiles.GetAll().Select(p => $"• {p.FullName} — {p.Description}")) +
                         $"\n\nThư viện: {TemplateLibrary.All.Count} mẫu dựng sẵn + mẫu cá nhân.\nHoạt động 100% offline.\n" +
                         $"Phiên bản {typeof(Connect).Assembly.GetName().Version}");
                    break;
                default:
                    if (id.StartsWith("tpl_")) Templates.Create(_app, TemplateLibrary.Find(id.Substring(4)), Settings);
                    else if (id.StartsWith("utpl_"))
                    {
                        var files = Templates.UserTemplates().ToList();
                        if (int.TryParse(id.Substring(5), out var i) && i < files.Count) Templates.OpenUser(_app, files[i]);
                    }
                    break;
            }
        }

        // ---- dropDown ----

        static readonly float[] Sizes = { 13, 14 };
        static readonly float[] Spacings = { 1.0f, 1.15f, 1.3f, 1.5f };

        public int GetItemCount(IRibbonControl c) =>
            c.Id == "ddStandard" ? Profiles.All.Length + 1 : c.Id == "ddSize" ? Sizes.Length : Spacings.Length;

        public string GetItemID(IRibbonControl c, int index) => c.Id + "_" + index;

        public string GetItemLabel(IRibbonControl c, int index)
        {
            if (c.Id == "ddStandard")
                return index == 0 ? "Tự nhận diện" : StandardLabel(Profiles.All[index - 1]);
            if (c.Id == "ddSize") return $"Cỡ {Sizes[index]}";
            return Spacings[index].ToString("0.0#");
        }

        public static string StandardLabel(StandardKind k)
        {
            var p = Profiles.Get(k);
            return $"{p.ShortName} — {p.Description}";
        }

        public int GetSelectedItemIndex(IRibbonControl c)
        {
            var s = Settings;
            if (c.Id == "ddStandard") return s.AutoDetect ? 0 : (int)s.Standard + 1;
            if (c.Id == "ddSize") return Math.Max(0, Array.IndexOf(Sizes, s.BodySize));
            int i = Array.FindIndex(Spacings, x => Math.Abs(x - s.LineSpacing) < 0.01);
            return i < 0 ? 1 : i;
        }

        public void OnSelect(IRibbonControl c, string selectedId, int index) => Safe(() =>
        {
            var s = Settings;
            if (c.Id == "ddStandard")
            {
                s.AutoDetect = index == 0;
                if (index > 0) s.Standard = Profiles.All[index - 1];
            }
            else if (c.Id == "ddSize") s.BodySize = Sizes[index];
            else s.LineSpacing = Spacings[index];
            s.Save();
        });

        // ---- toggle ----

        public bool GetPressed(IRibbonControl c)
        {
            try
            {
                if (c.Id == "tglPane") return CurrentPane()?.Visible ?? false;
                if (c.Id == "tglDuplex") return _app.Documents.Count > 0 && DocState.GetDuplex(_app.ActiveDocument);
            }
            catch { }
            return false;
        }

        public void OnToggle(IRibbonControl c, bool pressed) => Safe(() =>
        {
            if (c.Id == "tglPane") ShowPane(pressed);
            else if (c.Id == "tglDuplex")
            {
                var d = Doc();
                DocState.SetDuplex(d, pressed);
                foreach (Word.Section s in d.Sections) s.PageSetup.MirrorMargins = pressed ? -1 : 0;
                Status(pressed ? "In hai mặt: lề đối xứng trang chẵn/lẻ." : "In một mặt.");
            }
        });

        // ---- dynamicMenu ----

        public string GetTemplateMenu(IRibbonControl c)
        {
            var sb = new StringBuilder("<menu xmlns=\"http://schemas.microsoft.com/office/2009/07/customui\">");
            int g = 0;
            foreach (var grp in TemplateLibrary.All.GroupBy(t => t.Group))
            {
                sb.Append($"<menuSeparator id=\"sep{g++}\" title=\"{SecurityElement.Escape(grp.Key)}\"/>");
                foreach (var t in grp)
                    sb.Append($"<button id=\"tpl_{t.Id}\" label=\"{SecurityElement.Escape(t.Name)}\" onAction=\"OnAction\"/>");
            }
            var user = Templates.UserTemplates().ToList();
            sb.Append("<menuSeparator id=\"sepUser\" title=\"Mẫu cá nhân\"/>");
            if (user.Count == 0)
                sb.Append("<button id=\"btnTemplateFolder_m\" label=\"(Chưa có — mở thư mục mẫu)\" tag=\"btnTemplateFolder\" onAction=\"OnTagAction\"/>");
            for (int i = 0; i < user.Count; i++)
                sb.Append($"<button id=\"utpl_{i}\" label=\"{SecurityElement.Escape(Path.GetFileNameWithoutExtension(user[i]))}\" onAction=\"OnAction\"/>");
            sb.Append("</menu>");
            return sb.ToString();
        }

        public void OnTagAction(IRibbonControl control) => Safe(() => Execute(control.Tag));

        // ================= Tác vụ =================

        internal Word.Document Doc()
        {
            if (_app == null || _app.Documents.Count == 0)
                throw new InvalidOperationException("Chưa có tài liệu nào đang mở trong Microsoft Word.");

            try
            {
                if (_app.ActiveProtectedViewWindow != null)
                {
                    throw new InvalidOperationException(
                        "Tài liệu đang mở ở chế độ xem an toàn (Protected View).\n" +
                        "Vui lòng nhấn nút \"Enable Editing\" (Bật tính năng chỉnh sửa) phía trên Word trước khi tiếp tục.");
                }
            }
            catch (COMException) { }

            Word.Document doc;
            try
            {
                doc = _app.ActiveDocument;
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException("Không thể truy cập tài liệu hiện hành: " + ex.Message);
            }

            WordProtectionHelper.EnsureEditablePrintView(doc);
            return doc;
        }

        internal FormatResult RunFormat()
        {
            var r = Formatter.Run(_app, Doc(), Settings, Settings.AutoDetect ? (StandardKind?)null : Settings.Standard);
            Status(r.ToString());
            if (CurrentPane()?.Visible != true) Info(r.ToString());
            return r;
        }

        internal List<ReviewItem> RunReview()
        {
            var items = Review.Run(_app, Doc(), Settings, addComments: true);
            Pane()?.ShowResults("Rà soát thể thức", items);
            if (CurrentPane()?.Visible != true)
                Info(items.Count == 0 ? "Không phát hiện lỗi thể thức."
                    : $"Phát hiện {items.Count} lỗi thể thức, đã đánh dấu bằng nhận xét \"V-Format\".\n\n" +
                      string.Join("\n", items.Take(15).Select(x => "• " + x)) + (items.Count > 15 ? "\n..." : ""));
            return items;
        }

        internal List<ReviewItem> RunSpell(bool fix)
        {
            var items = Spell.Run(_app, Doc(), fix);
            Pane()?.ShowResults(fix ? "Đã sửa chính tả" : "Chính tả", items);
            if (CurrentPane()?.Visible != true)
                Info(items.Count == 0 ? "Không phát hiện lỗi chính tả." :
                    (fix ? $"Đã sửa {items.Count} lỗi." : $"Phát hiện {items.Count} lỗi (tô vàng + nhận xét)."));
            return items;
        }

        // ---- Cổng cho script kiểm thử (không hiện hộp thoại) ----

        /// <param name="standard">auto | nd30 | tt68 | ca08 | hd05 | hd29 | nd250 | cv5512</param>
        public string FormatActive(string standard)
        {
            StandardKind? k = null;
            if (string.Equals(standard, "auto", StringComparison.OrdinalIgnoreCase))
            {
                // "auto" luôn tự nhận diện, kể cả khi Ribbon đang chọn cố định một quy chuẩn
                var s = Settings; bool saved = s.AutoDetect;
                try { s.AutoDetect = true; var r0 = Formatter.Run(_app, Doc(), s, null); _testKind = r0.Profile.Kind; return r0.ToString(); }
                finally { s.AutoDetect = saved; }
            }
            else
                k = Profiles.All.First(x => string.Equals(Profiles.Get(x).ShortName.Replace(" ", "").Replace(".", ""),
                        standard.Replace(" ", ""), StringComparison.OrdinalIgnoreCase) || CodeOf(x) == standard.ToLowerInvariant());
            var r = Formatter.Run(_app, Doc(), Settings, k);
            _testKind = r.Profile.Kind;
            return r.ToString();
        }

        StandardKind? _testKind;

        /// <summary>Rà soát theo đúng quy chuẩn lần FormatActive gần nhất (không đổi cài đặt đã lưu).</summary>
        public string ReviewActive()
        {
            var s = Settings;
            bool auto = s.AutoDetect; var std = s.Standard;
            try
            {
                if (_testKind.HasValue) { s.AutoDetect = false; s.Standard = _testKind.Value; }
                return string.Join("\n", Review.Run(_app, Doc(), s, addComments: false).Select(x => x.ToString()));
            }
            finally { s.AutoDetect = auto; s.Standard = std; }
        }

        static string CodeOf(StandardKind k)
        {
            switch (k)
            {
                case StandardKind.HanhChinh_ND30: return "nd30";
                case StandardKind.QuanDoi_TT68: return "tt68";
                case StandardKind.CongAn_BCA08: return "ca08";
                case StandardKind.Dang_HD05: return "hd05";
                case StandardKind.Doan_HD29: return "hd29";
                case StandardKind.PhapDien_ND250: return "nd250";
                default: return "cv5512";
            }
        }

        void RunFormatTablesOnly()
        {
            var d = Doc();
            int n = 0;
            using (new WordBatch(_app, "V-Format: Chuẩn hóa bảng"))
                foreach (Word.Table t in d.Tables)
                {
                    try
                    {
                        if (t.Borders.Enable == 0) continue; // bảng khối đầu / chữ ký
                        t.Rows.LeftIndent = 0;
                        t.PreferredWidthType = Word.WdPreferredWidthType.wdPreferredWidthPercent;
                        t.PreferredWidth = 100;
                        t.AutoFitBehavior(Word.WdAutoFitBehavior.wdAutoFitWindow);
                        if (t.Rows.Count > 2) t.Rows[1].HeadingFormat = -1;
                        n++;
                    }
                    catch { }
                }
            Status($"Đã chuẩn hóa {n} bảng (vừa khổ giấy, lặp hàng tiêu đề).");
        }

        void OpenDictionary()
        {
            Directory.CreateDirectory(Settings.Folder);
            if (!File.Exists(Settings.UserDictionary))
                File.WriteAllText(Settings.UserDictionary,
                    "# Từ điển riêng V-Format: mỗi dòng một quy tắc  sai=đúng\r\n# Ví dụ:\r\nko=không\r\nUBND Tỉnh=UBND tỉnh\r\n",
                    new UTF8Encoding(true));
            Process.Start("notepad.exe", "\"" + Settings.UserDictionary + "\"");
        }

        void SaveTemplate()
        {
            var d = Doc();
            var name = Microsoft.VisualBasic.Interaction.InputBox("Tên mẫu cá nhân:", Title,
                Path.GetFileNameWithoutExtension(d.Name));
            if (string.IsNullOrWhiteSpace(name)) return;
            var path = Templates.SaveAsUserTemplate(d, name.Trim());
            _ribbon?.InvalidateControl("mnuTemplates");
            Info("Đã lưu vào kho mẫu cá nhân:\n" + path);
        }

        // ================= Task pane =================

        long WindowKey()
        {
            try { return _app.ActiveWindow.Hwnd; } catch { return 0; }
        }

        CustomTaskPane CurrentPane()
        {
            if (_app == null || _app.Windows.Count == 0) return null;
            _panes.TryGetValue(WindowKey(), out var p);
            return p;
        }

        TaskPaneHost Pane() => CurrentPane()?.ContentControl as TaskPaneHost;

        internal void ShowPane(bool visible)
        {
            if (_ctpFactory == null) throw new InvalidOperationException("Word chưa cung cấp Task Pane.");
            if (_app.Windows.Count == 0) throw new InvalidOperationException("Hãy mở một tài liệu trước.");
            var pane = CurrentPane();
            if (pane == null)
            {
                if (!visible) return;
                pane = _ctpFactory.CreateCTP(TaskPaneHost.ProgIdName, Title, _app.ActiveWindow);
                pane.DockPosition = MsoCTPDockPosition.msoCTPDockPositionRight;
                pane.Width = 340;
                ((TaskPaneHost)pane.ContentControl).Initialize(this);
                pane.VisibleStateChange += _ => _ribbon?.InvalidateControl("tglPane");
                _panes[WindowKey()] = pane;
            }
            pane.Visible = visible;
        }

        // ================= Tiện ích =================

        internal void Status(string msg)
        {
            try { _app.StatusBar = "V-FORMAT: " + msg.Replace("\n", "  "); } catch { }
            Pane()?.SetStatus(msg);
        }

        internal static void Info(string msg) =>
            MessageBox.Show(msg, Title, MessageBoxButtons.OK, MessageBoxIcon.Information);

        internal static bool Confirm(string msg) =>
            MessageBox.Show(msg, Title, MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes;

        internal void Safe(Action a)
        {
            try { a(); }
            catch (Exception ex)
            {
                try { _app.ScreenUpdating = true; } catch { }
                var e = ex is System.Reflection.TargetInvocationException t && t.InnerException != null ? t.InnerException : ex;
                MessageBox.Show(e.Message + (e is InvalidOperationException ? "" : "\n\n" + e.GetType().Name + "\n" + e.StackTrace),
                    Title, MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
            finally { _ribbon?.Invalidate(); }
        }

        void InsertTextAtSelection(string text, bool bold)
        {
            var sel = _app.Selection;
            if (sel == null) return;
            sel.Text = text;
            if (bold) sel.Font.Bold = 1;
            sel.Font.Name = "Times New Roman";
        }

        void InsertStampAtSelection(string stamp)
        {
            var sel = _app.Selection;
            if (sel == null) return;
            sel.Text = stamp;
            sel.Font.Bold = 1;
            sel.Font.Name = "Times New Roman";
            sel.Font.Size = 12;
            try { sel.Borders.Enable = 1; } catch { }
        }

    }
}
