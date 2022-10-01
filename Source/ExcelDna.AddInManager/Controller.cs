using ExcelDna.Integration;

namespace ExcelDna.AddInManager
{
    internal class Controller
    {
        public static void AutoOpen()
        {
            if (!Directory.Exists(installedDir))
                return;

            foreach (string xll in Directory.GetFiles(installedDir, "*.del"))
            {
                try
                {
                    File.Delete(xll);
                }
                catch
                {
                }
            }

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

        public static void Uninstall(string xllFileName)
        {
            string installedXllPath = Path.Combine(installedDir, Path.GetFileName(xllFileName));
            Unregister(installedXllPath);

            string delXllPath = installedXllPath + ".del";
            File.Move(installedXllPath, delXllPath, true);
        }

        private static void Register(string xllPath)
        {
            ExcelAsyncUtil.QueueAsMacro(() =>
            {
                ExcelIntegration.RegisterXLL(xllPath);
            });
        }

        private static void Unregister(string xllPath)
        {
            ExcelAsyncUtil.QueueAsMacro(() =>
            {
                ExcelIntegration.UnregisterXLL(xllPath);
            });
        }

        private static string dataDir = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "ExcelDna.AddInManager");
        private static string installedDir = Path.Combine(dataDir, "Installed");
    }
}
