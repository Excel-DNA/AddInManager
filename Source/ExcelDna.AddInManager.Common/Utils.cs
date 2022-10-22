using System.Diagnostics;

namespace ExcelDna.AddInManager.Common
{
    public class Utils
    {
        public const string IndexFileName = "index.xml";

        public static List<AddInFile> GetSourceAddins(string sourcePath)
        {
            List<AddInFile> addins = new();
            if (Directory.Exists(sourcePath))
                addins.AddRange(Directory.GetFiles(sourcePath, "*.xll").Select(i => GetAddInInfo(i)));

            return addins;
        }

        public static AddInFile GetAddInInfo(string path)
        {
            AddInFile result = new();
            result.FileName = Path.GetFileName(path);

            FileVersionInfo version = FileVersionInfo.GetVersionInfo(path);
            result.CompanyName = version.CompanyName;
            result.ProductName = version.ProductName;
            result.Version = version.FileVersion;

            if (result.ProductName == "Excel-DNA Add-In Framework for Microsoft Excel" && result.CompanyName == "Govert van Drimmelen")
            {
                result.CompanyName = null;
                result.ProductName = null;
                result.Version = null;
            }

            if (TryFindBitness(path, out Bitness bitness))
                result.Bitness = bitness;

            return result;
        }

        private static bool TryFindBitness(string exePath, out Bitness bitness)
        {
            bitness = Bitness.Unknown;

            try
            {
                using (var fileStream = File.OpenRead(exePath))
                {
                    using (var reader = new BinaryReader(fileStream))
                    {
                        // See http://www.microsoft.com/whdc/system/platform/firmware/PECOFF.mspx
                        // Offset to PE header is always at 0x3C.
                        // The PE header starts with "PE\0\0" =  0x50 0x45 0x00 0x00,
                        // followed by a 2-byte machine type field (see the document above for the enum).

                        fileStream.Seek(0x3c, SeekOrigin.Begin);
                        var peOffset = reader.ReadInt32();

                        fileStream.Seek(peOffset, SeekOrigin.Begin);
                        var peHead = reader.ReadUInt32();

                        if (peHead != 0x00004550) // "PE\0\0", little-endian
                        {
                            return false;
                        }

                        var machineType = (MachineType)reader.ReadUInt16();

                        switch (machineType)
                        {
                            case MachineType.ImageFileMachineI386:
                                {
                                    bitness = Bitness.Bit32;
                                    return true;
                                }
                            case MachineType.ImageFileMachineAmd64:
                            case MachineType.ImageFileMachineIa64:
                                {
                                    bitness = Bitness.Bit64;
                                    return true;
                                }
                            default:
                                {
                                    bitness = Bitness.Unknown;
                                    return false;
                                }
                        }
                    }
                }
            }
            catch
            {
            }

            return false;
        }

        private enum MachineType : ushort
        {
            ImageFileMachineAmd64 = 0x8664,
            ImageFileMachineI386 = 0x14c,
            ImageFileMachineIa64 = 0x200,
        }
    }
}
