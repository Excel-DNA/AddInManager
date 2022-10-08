using System.IO;

namespace ExcelDna.AddInManager
{
    internal partial class AddinsListView : System.Windows.Controls.UserControl
    {
        private class ListItem
        {
            public ListItem(AddInVersionInfo addin)
            {
                this.Addin = addin;
            }

            public string? CompanyName => Addin.CompanyName;
            public string? ProductName => Addin.IsVersioned ? Addin.ProductName : Path.GetFileNameWithoutExtension(Addin.Path);
            public string? Version => Addin.Version?.ToString();

            public AddInVersionInfo Addin { get; }
        }

        public AddinsListView()
        {
            InitializeComponent();
        }

        public void Add(List<AddInVersionInfo> addins)
        {
            foreach (var i in addins)
                addinsListView.Items.Add(new ListItem(i));
        }

        public List<AddInVersionInfo> GetSelectedAddins()
        {
            return addinsListView.SelectedItems.Cast<ListItem>().Select(i => i.Addin).ToList();
        }
    }
}
