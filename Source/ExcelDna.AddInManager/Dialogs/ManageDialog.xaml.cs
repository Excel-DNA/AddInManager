using System.Windows;

namespace ExcelDna.AddInManager
{
    internal partial class ManageDialog : Window
    {
        public ManageDialog(List<AddInVersionInfo> addins)
        {
            InitializeComponent();

            addinsListView.Add(addins);
        }

        public List<AddInVersionInfo>? GetAddinsForUninstall()
        {
            return addinsForUninstall;
        }

        private void OnUninstall(object sender, RoutedEventArgs args)
        {
            try
            {
                addinsForUninstall = addinsListView.GetSelectedAddins();

                DialogResult = true;
            }
            catch (Exception e)
            {
                ExceptionHandler.ShowException(e);
            }
        }

        private List<AddInVersionInfo>? addinsForUninstall;
    }
}
