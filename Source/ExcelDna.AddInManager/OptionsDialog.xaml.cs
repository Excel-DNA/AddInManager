using System.Windows;

namespace ExcelDna.AddInManager
{
    public partial class OptionsDialog : Window
    {
        public OptionsDialog(GeneralOptions options)
        {
            InitializeComponent();

            this.options = options;
            sourceTextBox.Text = options.source;
        }

        private void OnSave(object sender, RoutedEventArgs args)
        {
            try
            {
                options.source = sourceTextBox.Text;
                DialogResult = true;
            }
            catch (System.Exception e)
            {
                ExceptionHandler.ShowException(e);
            }
        }

        private GeneralOptions options;
    }
}
