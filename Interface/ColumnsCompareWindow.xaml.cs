using System.Collections.Generic;
using System.Windows;
using System.Windows.Controls;
using Converter.Models;
using Converter.Template_workbooks;
using ExcelRLibrary.TemplateWorkbooks;
using Telerik.Windows.Controls;

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
            viewModel = new CompareViewModel(CompareViewModel.DitctToObservDict(rulesDictionary), wsInfos, XlTemplateWorkbookType.LandProperty);
            DataContext = viewModel;
        }

        private void ListBox_OnSelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            viewModel.LastSelectedItem = ((RadListBox) sender).SelectedItem as string;
            viewModel.UpdateValuesExamples();
        }

        private void StartButton_OnClick(object sender, RoutedEventArgs e)
        {
            viewModel.CombineWorkbooks();
            Close();
        }


        private void AddColumnButtton_Click(object sender, RoutedEventArgs e)
        {
            var w = new EnterNameWindow();
            w.ShowDialog();
            var name = w.Name;
            viewModel.AddNewcolumn(name);
        }
    }
}
