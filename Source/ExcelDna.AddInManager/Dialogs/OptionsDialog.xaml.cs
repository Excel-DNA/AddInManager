using System.Windows;

namespace ExcelDna.AddInManager
{
    internal partial class OptionsDialog : Window
    {
        public OptionsDialog(GeneralOptions options)
        {
            InitializeComponent();

            this.options = options;

            autoUpdateAddInsCheckBox.IsChecked = options.autoUpdateAddIns;
            if (options.sources != null)
            {
                string text = "";
                foreach (var i in options.sources)
                {
                    if (text.Length > 0)
                        text += Environment.NewLine;
                    text += i.source;
                }
                sourcesTextBox.Text = text;
            }
        }

        private void OnSave(object sender, RoutedEventArgs args)
        {
            try
            {
                options.autoUpdateAddIns = autoUpdateAddInsCheckBox.IsChecked.GetValueOrDefault();
                List<AddInsSource> sources = new();
                for (int i = 0; i < sourcesTextBox.LineCount; ++i)
                {
                    AddInsSource source = new AddInsSource();
                    source.source = sourcesTextBox.GetLineText(i).Trim();
                    if (source.source.Length > 0)
                        sources.Add(source);
                }
                options.sources = sources;

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
