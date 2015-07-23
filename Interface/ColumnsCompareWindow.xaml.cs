using System.Collections.Generic;
using System.Threading.Tasks;
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

        public bool UseBase { set { viewModel.UseBase = value; } }
        public string BasePath { set { viewModel.BasePath = value; } }

        public ColumnsCompareWindow(Dictionary<JustColumn, List<string>> rulesDictionary, ICollection<WorksheetInfo> wsInfos, XlTemplateWorkbookType wbType)
        {
            InitializeComponent();
            this.wsInfos = wsInfos;
            viewModel = new CompareViewModel(CompareViewModel.DitctToObservDict(rulesDictionary), wsInfos, wbType);
            DataContext = viewModel;
        }

        private void ListBox_OnSelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            viewModel.LastSelectedItem = ((RadListBox) sender).SelectedItem as string;
            viewModel.UpdateValuesExamples();
        }

        private async void StartButton_OnClick(object sender, RoutedEventArgs e)
        {
            viewModel.WorkInProgress = true;
            viewModel.EditMode = false;
            await Task.Run(() => viewModel.CombineWorkbooks());
            viewModel.WorkInProgress = false;
            viewModel.EditMode = true;
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
