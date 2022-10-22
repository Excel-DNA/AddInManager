using ExcelDna.AddInManager.Common;

namespace ExcelDna.AddInManager.IndexGenerator
{
    internal class Program
    {
        static void Main(string[] args)
        {
            if (args.Length != 1 || !Directory.Exists(args[0]))
            {
                Console.WriteLine("Usage: ExcelDna.AddInManager.IndexGenerator.exe [Add-Ins Source Path]");
                return;
            }

            string sourcePath = args[0];
            List<AddInFile> addins = Utils.GetSourceAddins(sourcePath);

            string indexFile = Path.Combine(sourcePath, Utils.IndexFileName);
            XmlSerializer.XmlSerialize(indexFile, addins);

            Console.WriteLine("Generated " + indexFile);
        }
    }
}