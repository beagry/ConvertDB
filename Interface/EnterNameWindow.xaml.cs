using System.Windows;

namespace UI
{
    /// <summary>
    /// Interaction logic for EnterNameWindow.xaml
    /// </summary>
    public partial class EnterNameWindow : Window
    {
        public EnterNameWindow()
        {
            InitializeComponent();
        }

        public new string Name { get { return this.TextBox.Text; } }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }
    }
}
