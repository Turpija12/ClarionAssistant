using System;
using System.IO;
using System.Windows.Forms;

namespace TpsPreviewer
{
    internal static class Program
    {
        [STAThread]
        private static void Main(string[] args)
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);

            string initialFolder = null;
            if (args != null && args.Length > 0 && !string.IsNullOrEmpty(args[0]))
            {
                if (Directory.Exists(args[0]))
                    initialFolder = args[0];
                else if (File.Exists(args[0]))
                    initialFolder = Path.GetDirectoryName(args[0]);
            }

            Application.Run(new TpsPreviewForm(initialFolder));
        }
    }
}
