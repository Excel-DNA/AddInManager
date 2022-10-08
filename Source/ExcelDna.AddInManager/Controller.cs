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

        public void OnInstall()
        {
            InstallDialog dialog = new InstallDialog(GetSourceAddins());
            if (dialog.ShowDialog().GetValueOrDefault())
            {
                foreach (var i in dialog.GetSelectedAddins()!)
                {
                    Install(i);
                }
            }
        }

        public void OnManage()
        {
            ManageDialog dialog = new ManageDialog(GetInstalledAddins());
            if (dialog.ShowDialog().GetValueOrDefault())
            {
                foreach (var i in dialog.GetAddinsForUninstall()!)
                {
                    Uninstall(i.Path);
                }
            }
        }

        public void OnOptions()
        {
            OptionsDialog dialog = new OptionsDialog(generalOptions);
            if (dialog.ShowDialog().GetValueOrDefault())
            {
                Storage.SaveGeneralOptions(generalOptions);
            }
        }

        private static void Install(AddInVersionInfo addin)
        {
            if (addin.IsVersioned)
            {
                foreach (var i in GetInstalledAddins().Where(i => i.IsVersioned && i.CompanyName == addin.CompanyName && i.ProductName == addin.ProductName))
                {
                    Uninstall(i.Path);
                }
            }

            string installedXllPath = Path.Combine(installedDir, Path.GetFileName(addin.Path));
            if (File.Exists(installedXllPath))
            {
                Uninstall(installedXllPath);
            }
            Storage.CreateDirectoryForFile(installedXllPath);
            File.Copy(addin.Path, installedXllPath, true);
            Register(installedXllPath);
        }

        private static void Uninstall(string xllFileName)
        {
            string installedXllPath = Path.Combine(installedDir, xllFileName);
            Unregister(installedXllPath);

            string delXllPath = installedXllPath + ".del";
            File.Move(installedXllPath, delXllPath, true);
        }

        private static List<AddInVersionInfo> GetInstalledAddins()
        {
            List<AddInVersionInfo> addins = new();
            if (Directory.Exists(installedDir))
            {
                addins = Directory.GetFiles(installedDir, "*.xll").Select(i => new AddInVersionInfo(i)).ToList();
            }

            return addins;
        }

        private List<AddInVersionInfo> GetSourceAddins()
        {
            List<AddInVersionInfo> addins = new();
            if (generalOptions.sources != null)
            {
                foreach (var addinSource in generalOptions.sources)
                {
                    string? source = addinSource.source;
                    if (!string.IsNullOrWhiteSpace(source) && Directory.Exists(source))
                    {
                        addins.AddRange(Directory.GetFiles(source, "*.xll").Select(i => new AddInVersionInfo(i)));
                    }
                }
            }

            return addins;
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
