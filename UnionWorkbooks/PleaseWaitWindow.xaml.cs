using System.Windows;

namespace UnionWorkbooks
{
    /// <summary>
    /// Interaction logic for PleaseWaitWindow.xaml
    /// </summary>
    public partial class PleaseWaitWindow : Window
    {
        public PleaseWaitWindow()
        {
            InitializeComponent();
        }


        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            double screenWidth = SystemParameters.PrimaryScreenWidth;
            double screenHeight = SystemParameters.PrimaryScreenHeight;
            double windowWidth = this.Width;
            double windowHeight = this.Height;
            this.Left = (screenWidth / 2) - (windowWidth / 2);
            this.Top = (screenHeight / 2) - (windowHeight / 2);
        }
    }
}
