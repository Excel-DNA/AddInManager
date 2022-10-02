using ExcelDna.Integration;
using System.Diagnostics.CodeAnalysis;
using System.IO;

namespace ExcelDna.AddInManager
{
    internal class Controller
    {
        public Controller()
        {
            LoadGeneralOptions();

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
            string installedXllPath = Path.Combine(installedDir, Path.GetFileName(sourceXllPath));
            Storage.CreateDirectoryForFile(installedXllPath);
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

        public void OnOptions()
        {
            OptionsDialog dialog = new OptionsDialog(generalOptions);
            if (dialog.ShowDialog().GetValueOrDefault())
            {
                Storage.SaveGeneralOptions(generalOptions);
            }
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

        [MemberNotNull(nameof(generalOptions))]
        private void LoadGeneralOptions()
        {
            try
            {
                generalOptions = Storage.LoadGeneralOptions() ?? null!;
            }
            catch (System.ApplicationException e)
            {
                System.Windows.MessageBox.Show(e.ToString(), "ExcelDna.AddInManager Load general options", System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Error);
            }
            if (generalOptions == null)
                generalOptions = new GeneralOptions();
        }

        private static string installedDir = Storage.GetInstalledAddinsDirectory();
        private GeneralOptions generalOptions;
    }
}
