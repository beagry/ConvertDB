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

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            var w = new ConverterWindow {Owner = this};
            this.Hide();
            w.Show();
        }

        private void Button_Click_1(object sender, RoutedEventArgs e)
        {
            var w = new ExtentionConverterWindow { Owner = this };
            this.Hide();
            w.Show();
        }
    }
}
