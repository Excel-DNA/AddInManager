using ExcelDna.AddInManager.Common;

namespace ExcelDna.AddInManager
{
    internal class AddInVersionInfo
    {
        public AddInVersionInfo(string path) : this(Utils.GetAddInInfo(path), path)
        {
        }

        public AddInVersionInfo(string sourceDirectory, AddInFile addInFile) : this(addInFile, System.IO.Path.Combine(sourceDirectory, addInFile.FileName ?? string.Empty))
        {
        }

        private AddInVersionInfo(AddInFile addInFile, string path)
        {
            Path = path;
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
