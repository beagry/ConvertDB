using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;
using Converter;
using RwaySupportLibraly;

namespace Interface
{
    /// <summary>
    /// Interaction logic for ConverterWindow.xaml
    /// </summary>
    public partial class ConverterWindow
    {
        private CoverterArgs _coverterArgs;
        public ConverterWindow()
        {
            InitializeComponent();
            _coverterArgs = new CoverterArgs();

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
