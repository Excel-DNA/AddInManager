using ExcelDna.AddInManager.Common;
using System.IO;

namespace ExcelDna.AddInManager
{
    internal class Storage
    {
        /// <exception cref="ApplicationException"></exception>
        public static void SaveGeneralOptions(GeneralOptions options)
        {
            string file = GetGeneralOptionsFile();
            CreateDirectoryForFile(file);
            XmlSerializer.XmlSerialize(file, options);
        }

        /// <exception cref="ApplicationException"></exception>
        public static GeneralOptions? LoadGeneralOptions()
        {
            string file = GetGeneralOptionsFile();
            if (!File.Exists(file))
                return null;

            return XmlSerializer.XmlDeserialize<GeneralOptions>(file);
        }

        public static void CreateDirectoryForFile(string file)
        {
            string directory = Path.GetDirectoryName(file)!;
            if (!Directory.Exists(directory))
                Directory.CreateDirectory(directory);
        }

        public static string GetInstalledAddinsDirectory()
        {
            return GetAppDataFile("Installed");
        }

        private static string GetGeneralOptionsFile()
        {
            return GetAppDataFile("GeneralOptions.xml");
        }

        private static string GetAppDataFile(string file)
        {
            return Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), @"ExcelDna.AddInManager", file);
        }
    }
}
