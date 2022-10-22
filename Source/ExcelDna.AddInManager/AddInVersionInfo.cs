using ExcelDna.AddInManager.Common;

namespace ExcelDna.AddInManager
{
    internal class AddInVersionInfo
    {
        public AddInVersionInfo(string path)
        {
            Path = path;

            AddInFile addInFile = Utils.GetAddInInfo(path);
            CompanyName = addInFile.CompanyName;
            ProductName = addInFile.ProductName;
            Bitness = addInFile.Bitness;
            if (Version.TryParse(addInFile.Version, out Version? version))
                Version = version;

            IsVersioned = !string.IsNullOrWhiteSpace(CompanyName) && !string.IsNullOrWhiteSpace(ProductName) && Version != null;
        }

        public string Path { get; }
        public bool IsVersioned { get; }
        public string? CompanyName { get; }
        public string? ProductName { get; }
        public Version? Version { get; }
        public Bitness Bitness { get; }
    }
}
