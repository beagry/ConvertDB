using System.Windows;

namespace UI
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        private void ConvertButton_Click(object sender, RoutedEventArgs e)
        {
            var w = new ConverterWindow();
            w.Show();
            this.Close();
        }

        private void ChangeExtentionButton_Click(object sender, RoutedEventArgs e)
        {
            var w = new ExtentionConverterWindow();
            w.Show();
            this.Close();
        }
    }
}
