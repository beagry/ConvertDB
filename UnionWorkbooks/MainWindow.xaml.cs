using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Input;
using System.Windows.Threading;
using Converter;
using ExcelRLibrary;
using Telerik.Windows.Controls;
using Excel = Microsoft.Office.Interop.Excel;

namespace UnionWorkbooks
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private ViewModel viewModel;

        public Excel.Application ExcelApp { get { return ExcelHelper.App; } }

        public MainWindow()
        {
            InitializeComponent();
            ResetParams();
            DataContext = viewModel;

            ExcelHelper.App = ExcelHelper.CreateNewApplication();

            viewModel.WorksheetsToCopy.CollectionChanged += WorksheetsToCopy_CollectionChanged;
        }

        void WorksheetsToCopy_CollectionChanged(object sender, System.Collections.Specialized.NotifyCollectionChangedEventArgs e)
        {
            var coll = sender as ObservableCollection<string>;
            if (coll == null) return;

            foreach (var workbook in viewModel.Workbooks)
                workbook.WorksheetsForCountMaxRows = new List<string>(coll);

            UpdateTotalItems();
        }

        private void ListBox_Drop(object sender, DragEventArgs e)
        {
            var files = (string[])e.Data.GetData(DataFormats.FileDrop);

            if (files == null) return;
            
            files.ToList().ForEach(s =>
            {
                if (viewModel.Workbooks.All(w => w.Path != s && FileTypeChecker.IsFileExtelType(s)))
                {
                    var newWB = new WorkbookWithItemsQnt(s);
                    viewModel.Workbooks.Add(newWB);
                    newWB.WorksheetsForCountMaxRows = new List<string>(viewModel.WorksheetsToCopy);
                }
            });

            ManualUpdateWindow();
        }

        private void ConverterWindow_OnKeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Delete)
            {
                var focusetControl = FocusManager.GetFocusedElement(this);
                if (focusetControl == null) return;

                if (focusetControl is ListBoxItem)
                {
                    if (Equals(((ListBoxItem) focusetControl).GetVisualParent<ListBox>(), WorkbooksListBox))
                    {
                        var selItems = WorkbooksListBox.SelectedItems;
                        if (selItems == null || selItems.Count == 0) return;

                        foreach (var item in selItems.Cast<WorkbookWithItemsQnt>().ToList())
                        {
                            viewModel.Workbooks.Remove(item);
                        }
                        ManualUpdateWindow();
                    }
                }
                else if (focusetControl is RadListBoxItem)
                    if (Equals(((RadListBoxItem) focusetControl).GetVisualParent<RadListBox>(),SelectedWorksheetsListBox))
                    {
                        var selItems = SelectedWorksheetsListBox.SelectedItems;
                        if (selItems == null || selItems.Count == 0) return;

                        foreach (var item in selItems.Cast<string>().ToList())
                        {
                            viewModel.WorksheetsToCopy.Remove(item);
                        }
                        ManualUpdateWindow();
                    }
            }
            else if (e.Key == Key.Enter)
            {
                StartCombine();
            }
            
        }

        private async void ButtonBase_OnClick(object sender, RoutedEventArgs e)
        {
//            await CombineAsync();
            BlockUi();
            StartCombine();
            UnblockUi();
        }

        private async Task CombineAsync()
        {
            await Task.Factory.StartNew(() =>
            {
                BlockUi();
                StartCombine();
                UnblockUi();
            });

        }


        private void StartCombine()
        {
            if (viewModel.WorksheetsToCopy.Count == 0) return;

            var resultWb = ExcelHelper.CreateNewWorkbook(ExcelHelper.App, (byte)viewModel.WorksheetsToCopy.Count());
            var sampleWb = ExcelHelper.GetWorkbook(ExcelApp,viewModel.Workbooks.First().Path);
            for (int i = 1; i <= viewModel.WorksheetsToCopy.Count; i++)
            {
                var resultWs = (Excel.Worksheet)resultWb.Worksheets[i];
                resultWs.Name = viewModel.WorksheetsToCopy[i - 1];

                var sourceWs =
                    sampleWb.Worksheets.Cast<Excel.Worksheet>()
                        .FirstOrDefault(w => string.Equals(resultWs.Name, w.Name, StringComparison.OrdinalIgnoreCase));

                WriteWideHead(resultWs, sourceWs, viewModel.HeadSize);
            }

            var fillers =
                resultWb.Worksheets.Cast<Excel.Worksheet>().Select(w => new WorksheetFiller(w)).ToList();


            MyProgressBar.Maximum = viewModel.Workbooks.Count;

            foreach (var workbookInfo in viewModel.Workbooks)
            {
                var wb = ExcelHelper.GetWorkbook(ExcelApp,workbookInfo.Path);

                foreach (var targetWs in resultWb.Worksheets.Cast<Excel.Worksheet>())
                {
                    var sourceWs =
                        wb.Worksheets.Cast<Excel.Worksheet>()
                            .FirstOrDefault(
                                w => String.Equals(w.Name, targetWs.Name, StringComparison.OrdinalIgnoreCase));
                    if (sourceWs == null) continue;

                    var filler =
                        fillers.First(
                            f => string.Equals(f.WorksheetName, targetWs.Name, StringComparison.OrdinalIgnoreCase));

                    filler.InsertWorksheet(sourceWs, viewModel.HeadSize + 1);
                }

                wb.Close();

                MyProgressBar.Value ++;
                Thread.Sleep(1000);
            }

            ((Excel.Application)resultWb.Parent).Visible = true;
            resultWb.Activate();
            ((Excel.Worksheet)resultWb.Worksheets[1]).Activate();
            ResetParams();
        }

        private void BlockUi()
        {
            WorkbooksListBox.IsEnabled = false;
            SelectedWorksheetsListBox.IsEnabled = false;
            AllWorksheetsListBox.IsEnabled = false;
        }

        private void UnblockUi()
        {
            WorkbooksListBox.IsEnabled = true;
            SelectedWorksheetsListBox.IsEnabled = true;
            AllWorksheetsListBox.IsEnabled = true;
        }

        private void ResetParams()
        {
            viewModel = new ViewModel();
            WorkbooksListBox.ItemsSource = viewModel.Workbooks;
            SelectedWorksheetsListBox.ItemsSource = viewModel.WorksheetsToCopy;
            MyProgressBar.Value = 0;
            ManualUpdateWindow();
        }

        private void ManualUpdateWindow()
        {
            UpdateTotalItems();
            AllWorksheetsListBox.ItemsSource = viewModel.AllWorksheetsCollection;
        }

        private void UpdateTotalItems()
        {
            TotalItemsLabel.Content = viewModel.Workbooks.Sum(w => w.MaxRowsInWorkbook);
        }



        private void WriteWideHead(Excel.Worksheet targetWs, Excel.Worksheet soureWs, byte headSize)
        {
            for (int i = 1; i <= headSize; i++)
            {
                ((Excel.Range)targetWs.Rows[i]).EntireRow.Value2 = ((Excel.Range)soureWs.Rows[i]).EntireRow.Value2;
            }
        }
    }
}
