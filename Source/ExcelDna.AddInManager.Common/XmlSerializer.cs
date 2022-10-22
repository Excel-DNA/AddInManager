namespace ExcelDna.AddInManager.Common
{
    public class XmlSerializer
    {
        /// <exception cref="ApplicationException"></exception>
        public static void XmlSerialize<T>(string file, T o)
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
        public static T XmlDeserialize<T>(string file)
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
    }
}
