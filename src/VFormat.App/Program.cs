using System;
using System.Windows.Forms;

namespace VFormat.App
{
    static class Program
    {
        [STAThread]
        static void Main(string[] args)
        {
            ApplicationConfiguration.Initialize();
            var form = new MainForm();
            // kéo thả file lên VFormat.exe hoặc "Open with" → xử lý ngay khi mở
            if (args.Length > 0) form.Shown += (s, e) => form.Enqueue(args);
            Application.Run(form);
        }
    }
}
