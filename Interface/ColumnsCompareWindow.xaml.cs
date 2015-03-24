using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Windows;
using System.Windows.Controls;
using Converter;
using Telerik.Windows.Controls;

namespace UI
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
