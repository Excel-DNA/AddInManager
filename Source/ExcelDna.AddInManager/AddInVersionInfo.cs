using ExcelDna.AddInManager.Common;

namespace ExcelDna.AddInManager
{
    internal class AddInVersionInfo
    {
        public AddInVersionInfo(string path) : this(Utils.GetAddInInfo(path), path, null)
        {
        }

        public AddInVersionInfo(string sourceDirectory, AddInFile addInFile) : this(addInFile, System.IO.Path.Combine(sourceDirectory, addInFile.FileName ?? string.Empty), null)
        {
        }

        public AddInVersionInfo(Uri sourceDirectory, AddInFile addInFile) : this(addInFile, null, new Uri(sourceDirectory, addInFile.FileName))
        {
        }

        private AddInVersionInfo(AddInFile addInFile, string? path, Uri? uri)
        {
            Path = path;
            Uri = uri;
            CompanyName = addInFile.CompanyName;
            ProductName = addInFile.ProductName;
            Bitness = addInFile.Bitness;
            if (Version.TryParse(addInFile.Version, out Version? version))
                Version = version;

            IsVersioned = !string.IsNullOrWhiteSpace(CompanyName) && !string.IsNullOrWhiteSpace(ProductName) && Version != null;
        }

        public string? Path { get; }
        public Uri? Uri { get; }
        public bool IsVersioned { get; }
        public string? CompanyName { get; }
        public string? ProductName { get; }
        public Version? Version { get; }
        public Bitness Bitness { get; }
    }
}
