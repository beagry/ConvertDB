using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using Converter;
using Converter.Models;
using Converter.Template_workbooks;
using ExcelRLibrary;
using Telerik.Windows.Controls;

namespace UI
{
    /// <summary>
    /// Interaction logic for Window1.xaml
    /// </summary>
    public partial class ColumnsCompareWindow : Window
    {
        private CompareViewModel view;
        private ICollection<WorksheetInfo> wsInfos; 

        public ColumnsCompareWindow(Dictionary<string, List<string>> rulesDictionary, List<WorksheetInfo> wsInfos)
        {
            InitializeComponent();
            this.wsInfos = wsInfos;
            view = new CompareViewModel(DitctToObservDict(rulesDictionary), wsInfos);
            UnbindexListBox.ItemsSource = view.UnbindedColumns;
            BindedColumnsListBox.ItemsSource = view.BindedColumnsDictionary;
            ValuesExamplesListBox.ItemsSource = view.LastSelectedColumnValuesExamples;
        }

        private void ListBox_OnSelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            view.LastSelectedItem = ((RadListBox) sender).SelectedItem as string;
            ValuesExamplesListBox.ItemsSource = view.LastSelectedColumnValuesExamples;
        }

        private void StartButton_OnClick(object sender, RoutedEventArgs e)
        {
            var dict = ObservDictToDict(view.BindedColumnsDictionary);

            var typifer = new WorkbookTypifier<LandPropertyTemplateWorkbook>()
            {
                RulesDictionary = dict,
                WorkbooksPaths = wsInfos.Select(w => w.Workbook.Path).Distinct().ToList()
            };

            var result = typifer.CombineToSingleWorkbook();
            if (result == null) return;

            result.SaveWithDialog("Обработанная выгрузка");
            this.Close();
        }

        private Dictionary<string, ObservableCollection<string>> DitctToObservDict(
            Dictionary<string, List<string>> sourceDict)
        {
            return sourceDict.ToDictionary(k => k.Key, v => new ObservableCollection<string>(v.Value));
        }

        private Dictionary<string, List<string>> ObservDictToDict(
            Dictionary<string, ObservableCollection<string>> sourceDict)
        {
            return sourceDict.ToDictionary(k => k.Key, v => v.Value.ToList());
        }
    }
}
