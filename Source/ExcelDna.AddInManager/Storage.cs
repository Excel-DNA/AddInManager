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
            XmlSerialize(file, options);
        }

        /// <exception cref="ApplicationException"></exception>
        public static GeneralOptions? LoadGeneralOptions()
        {
            string file = GetGeneralOptionsFile();
            if (!File.Exists(file))
                return null;

            return XmlDeserialize<GeneralOptions>(file);
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

        /// <exception cref="ApplicationException"></exception>
        private static void XmlSerialize<T>(string file, T o)
        {
            try
            {
                System.Xml.XmlWriterSettings settings = new System.Xml.XmlWriterSettings();
                settings.Indent = true;
                using (System.Xml.XmlWriter stream = System.Xml.XmlWriter.Create(file, settings))
                {
                    CreateSerializer<T>().Serialize(stream, o);
                }
            }
            catch (Exception e)
            {
                throw new ApplicationException(e.Message);
            }
        }

        /// <exception cref="ApplicationException"></exception>
        private static T XmlDeserialize<T>(string file)
        {
            try
            {
                using (System.Xml.XmlReader stream = System.Xml.XmlReader.Create(file))
                {
                    return (T)CreateSerializer<T>().Deserialize(stream)!;
                }
            }
            catch (Exception e)
            {
                throw new ApplicationException(e.Message);
            }
        }

        private static System.Xml.Serialization.XmlSerializer CreateSerializer<T>()
        {
            return new System.Xml.Serialization.XmlSerializer(typeof(T));
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
