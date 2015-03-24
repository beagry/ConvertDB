using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Runtime.Remoting.Channels;
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
using Telerik.Windows.Controls;

namespace WpfApplication2
{
    /// <summary>
    /// Interaction logic for Window1.xaml
    /// </summary>
    public partial class ColumnsCompareWindow : Window
    {
        private CompareViewModel view;
        public ColumnsCompareWindow(Dictionary<string, ObservableCollection<string>> dict, ICollection<WorksheetInfo> wsInfos)
        {
            InitializeComponent();
            view = new CompareViewModel(dict,wsInfos);
//            this.DataContext = view;
            UnbindexListBox.ItemsSource = view.UnbindedColumns;
            BindedColumnsListBox.ItemsSource = view.BindedColumnsDictionary;
            ValuesExamplesListBox.ItemsSource = view.LastSelectedColumnValuesExamples;
        }

        private void ListBox_OnSelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            view.LastSelectedItem = ((RadListBox) sender).SelectedItem as string;
            ValuesExamplesListBox.ItemsSource = view.LastSelectedColumnValuesExamples;
        }
    }
}
