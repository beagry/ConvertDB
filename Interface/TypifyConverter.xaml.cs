using System;
using System.Windows;
using Converter;
using RwaySupportLibraly;

namespace UI
{
    /// <summary>
    /// Interaction logic for ConverterWindow.xaml
    /// </summary>
    public partial class ConverterWindow
    {
        private ConverterArgs _coverterArgs;
        public ConverterWindow()
        {
            InitializeComponent();
            _coverterArgs = new ConverterArgs();

            foreach (Enum e in Enum.GetValues(typeof (XlTemplateWorkbookTypes)))
                WorkbookTypesComboBox.Items.Add(e.GetDescription());
        }

        private void ListBox_Drop(object sender, DragEventArgs e)
        {
            WindowsExtentions.ListBox_DropWorkbook(sender, e);
        }

        private void Window_Closing(object sender, System.ComponentModel.CancelEventArgs e)
        {
            var window = sender as Window;
            if (window != null && window.Owner != null)
                window.Owner.Close();
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {

        }
    }
}
