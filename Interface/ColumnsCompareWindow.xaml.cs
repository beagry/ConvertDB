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
using JustColumn = ExcelRLibrary.TemplateWorkbooks.JustColumn;

namespace UI
{
    /// <summary>
    /// Interaction logic for Window1.xaml
    /// </summary>
    public partial class ColumnsCompareWindow : Window
    {
        private readonly CompareViewModel viewModel;
        private readonly ICollection<WorksheetInfo> wsInfos; 

        public ColumnsCompareWindow(Dictionary<JustColumn, List<string>> rulesDictionary, ICollection<WorksheetInfo> wsInfos)
        {
            InitializeComponent();
            this.wsInfos = wsInfos;
            viewModel = new CompareViewModel(DitctToObservDict(rulesDictionary), wsInfos);
            DataContext = viewModel;
        }

        private void ListBox_OnSelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            viewModel.LastSelectedItem = ((RadListBox) sender).SelectedItem as string;
            viewModel.UpdateValuesExamples();
        }

        private void StartButton_OnClick(object sender, RoutedEventArgs e)
        {
            var dict = ObservDictToDict(viewModel.BindedColumnsDictionary).ToDictionary(k => k.Key.CodeName,v => v.Value);

            var typifer = new WorkbookTypifier<LandPropertyTemplateWorkbook>()
            {
                RulesDictionary = dict,
                WorkbooksPaths = wsInfos.Select(w => w.Workbook.Path).Distinct().ToList()
            };

            var result = typifer.CombineToSingleWorkbook();
            if (result == null) return;

            result.SaveWithDialog("Обработанная выгрузка");
            Close();
        }

        private Dictionary<JustColumn, ObservableCollection<string>> DitctToObservDict(
            Dictionary<JustColumn, List<string>> sourceDict)
        {
            return sourceDict.ToDictionary(k => k.Key, v => new ObservableCollection<string>(v.Value));
        }

        private Dictionary<JustColumn, List<string>> ObservDictToDict(
            Dictionary<JustColumn, ObservableCollection<string>> sourceDict)
        {
            return sourceDict.ToDictionary(k => k.Key, v => v.Value.ToList());
        }
    }
}
