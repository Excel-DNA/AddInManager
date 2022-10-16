using System.Diagnostics;
using System.IO;

namespace ExcelDna.AddInManager
{
    internal class AddInVersionInfo
    {
        public AddInVersionInfo(string path)
        {
            Path = path;

            FileVersionInfo version = FileVersionInfo.GetVersionInfo(path);
            CompanyName = version.CompanyName;
            ProductName = version.ProductName;

            if (version.FileVersion != null)
            {
                Version = new Version(version.FileMajorPart, version.FileMinorPart, version.FileBuildPart, version.FilePrivatePart);
            }

            if (ProductName == "Excel-DNA Add-In Framework for Microsoft Excel" && CompanyName == "Govert van Drimmelen")
            {
                CompanyName = null;
                ProductName = null;
                Version = null;
            }

            IsVersioned = !string.IsNullOrWhiteSpace(CompanyName) && !string.IsNullOrWhiteSpace(ProductName) && Version != null;

            if (TryFindBitness(path, out Bitness bitness))
                Bitness = bitness;
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

        public string Path { get; }
        public bool IsVersioned { get; }
        public string? CompanyName { get; }
        public string? ProductName { get; }
        public Version? Version { get; }
        public Bitness Bitness { get; }
    }
}
