using System.Diagnostics;

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
        }

        public string Path { get; }
        public bool IsVersioned { get; }
        public string? CompanyName { get; }
        public string? ProductName { get; }
        public Version? Version { get; }
    }
}
