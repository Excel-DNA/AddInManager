using System.Windows;

namespace ExcelDna.AddInManager
{
    internal partial class InstallDialog : Window
    {
        public InstallDialog(List<AddInVersionInfo> addins)
        {
            InitializeComponent();

            addinsListView.Add(addins);
        }

        public List<AddInVersionInfo>? GetSelectedAddins()
        {
            return selectedAddins;
        }

        private void OnInstall(object sender, RoutedEventArgs args)
        {
            try
            {
                selectedAddins = addinsListView.GetSelectedAddins();

                DialogResult = true;
            }
            catch (Exception e)
            {
                ExceptionHandler.ShowException(e);
            }
        }

        private List<AddInVersionInfo>? selectedAddins;
    }
}
