using System.IO;
using System.Windows;

namespace ExcelDna.AddInManager
{
    public partial class InstallDialog : Window
    {
        public InstallDialog(List<string> files)
        {
            InitializeComponent();
            this.files = files;

            foreach (string i in files)
                addinsListBox.Items.Add(Path.GetFileNameWithoutExtension(i));
        }

        public List<string>? GetSelectedFiles()
        {
            return selectedFiles;
        }

        private void OnInstall(object sender, RoutedEventArgs args)
        {
            try
            {
                selectedFiles = new List<string>();
                int i = addinsListBox.SelectedIndex;
                if (i >= 0)
                    selectedFiles.Add(files[i]);

                DialogResult = true;
            }
            catch (Exception e)
            {
                ExceptionHandler.ShowException(e);
            }
        }

        private List<string> files;
        private List<string>? selectedFiles;
    }
}
