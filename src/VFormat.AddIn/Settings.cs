using System;
using System.IO;
using System.Xml.Serialization;
using VFormat.Core;

namespace VFormat.AddIn
{
    /// <summary>Thiết lập người dùng, lưu tại %AppData%\VFormat\settings.xml.</summary>
    public class Settings
    {
        public bool AutoDetect = true;
        public StandardKind Standard = StandardKind.HanhChinh_ND30;
        public float BodySize = 14;
        public float LineSpacing = 1.15f;
        public string DiaDanh = "Hà Nội";
        public bool FixText = true;
        public bool BuildTables = true;
        public bool DrawLines = true;
        public bool PageNumbers = true;
        public bool FormatTables = true;

        public static string Folder => Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "VFormat");
        public static string TemplateFolder => Path.Combine(Folder, "Templates");
        public static string UserDictionary => Path.Combine(Folder, "tu-dien.txt");
        static string FilePath => Path.Combine(Folder, "settings.xml");

        static Settings _current;
        public static Settings Current => _current ?? (_current = Load());

        public static event Action Changed;

        static Settings Load()
        {
            try
            {
                if (File.Exists(FilePath))
                    using (var f = File.OpenRead(FilePath))
                        return (Settings)new XmlSerializer(typeof(Settings)).Deserialize(f);
            }
            catch { /* tệp hỏng: dùng mặc định */ }
            return new Settings();
        }

        public void Save()
        {
            try
            {
                Directory.CreateDirectory(Folder);
                using (var f = File.Create(FilePath))
                    new XmlSerializer(typeof(Settings)).Serialize(f, this);
            }
            catch { }
            Changed?.Invoke();
        }

        /// <summary>Hồ sơ quy chuẩn áp dụng cho các đoạn đã đọc (tự nhận diện nếu bật).</summary>
        public FormatProfile ResolveProfile(System.Collections.Generic.IList<ParagraphInfo> paragraphs)
        {
            var kind = AutoDetect && paragraphs != null ? StandardDetector.Detect(paragraphs) : Standard;
            return ProfileFor(kind);
        }

        public FormatProfile ProfileFor(StandardKind kind)
        {
            var p = Profiles.Get(kind).Clone();
            p.SetBodySize(BodySize);
            p.SetLineSpacing(LineSpacing);
            return p;
        }
    }
}
