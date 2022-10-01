using ExcelDna.Integration;

namespace ExcelDna.AddInManager
{
    internal class Controller
    {
        public static void AutoOpen()
        {
            if (!Directory.Exists(installedDir))
                return;

            foreach (string xll in Directory.GetFiles(installedDir, "*.xll"))
            {
                Register(xll);
            }
        }

        public static void Install(string sourceXllPath)
        {
            Directory.CreateDirectory(installedDir);
            string installedXllPath = Path.Combine(installedDir, Path.GetFileName(sourceXllPath));
            File.Copy(sourceXllPath, installedXllPath, true);
            Register(installedXllPath);
        }

        private static void Register(string xllPath)
        {
            ExcelAsyncUtil.QueueAsMacro(() =>
            {
                ExcelIntegration.RegisterXLL(xllPath);
            });
        }

        private static string dataDir = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "ExcelDna.AddInManager");
        private static string installedDir = Path.Combine(dataDir, "Installed");
    }
}
