using System.IO;
using System.Windows;

namespace ExcelDna.AddInManager
{
    internal partial class InstallDialog : Window
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

        public InstallDialog(List<AddInVersionInfo> addins)
        {
            InitializeComponent();
            this.addins = addins;

            foreach (var i in addins)
                addinsListView.Items.Add(new ListItem(i));
        }

        public List<AddInVersionInfo>? GetSelectedAddins()
        {
            return selectedAddins;
        }

        private void OnInstall(object sender, RoutedEventArgs args)
        {
            try
            {
                selectedAddins = new List<AddInVersionInfo>();
                int i = addinsListView.SelectedIndex;
                if (i >= 0)
                    selectedAddins.Add(addins[i]);

                DialogResult = true;
            }
            catch (Exception e)
            {
                ExceptionHandler.ShowException(e);
            }
        }

        private List<AddInVersionInfo> addins;
        private List<AddInVersionInfo>? selectedAddins;
    }
}
