using System.IO;
using System.Windows;

namespace ExcelDna.AddInManager
{
    public partial class ManageDialog : Window
    {
        public ManageDialog(List<string> files)
        {
            InitializeComponent();
            this.files = files;

            foreach (string i in files)
                addinsListBox.Items.Add(Path.GetFileNameWithoutExtension(i));
        }

        public List<string>? GetFilesForUninstall()
        {
            return filesForUninstall;
        }

        private void OnUninstall(object sender, RoutedEventArgs args)
        {
            try
            {
                filesForUninstall = new List<string>();
                int i = addinsListBox.SelectedIndex;
                if (i >= 0)
                    filesForUninstall.Add(files[i]);

                DialogResult = true;
            }
            catch (Exception e)
            {
                ExceptionHandler.ShowException(e);
            }
        }

        private List<string> files;
        private List<string>? filesForUninstall;
    }
}
