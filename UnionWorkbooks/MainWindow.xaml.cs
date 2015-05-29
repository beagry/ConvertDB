using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Collections.Specialized;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Input;
using Converter;
using ExcelRLibrary;
using Microsoft.Office.Interop.Excel;
using Telerik.Windows.Controls;
using Application = Microsoft.Office.Interop.Excel.Application;
using ListBox = System.Windows.Controls.ListBox;
using TextBox = System.Windows.Controls.TextBox;
using Window = System.Windows.Window;

namespace UnionWorkbooks
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private ViewModel viewModel;

        public Application ExcelApp { get { return ExcelHelper.App; } }

        public MainWindow()
        {
            InitializeComponent();

            ResetParams();
            DataContext = viewModel;

            ExcelHelper.App = ExcelHelper.GetApplication();
        }

        void WorksheetsToCopy_CollectionChanged(object sender, NotifyCollectionChangedEventArgs e)
        {
            var coll = sender as ObservableCollection<string>;
            if (coll == null) return;

            foreach (var workbook in viewModel.Workbooks)
                workbook.WorksheetsForCountMaxRows = new List<string>(coll);
        }

        private void ListBox_Drop(object sender, DragEventArgs e)
        {
            var waitWindow = new PleaseWaitWindow(){Owner = this};
            waitWindow.Show();

            var files = (string[])e.Data.GetData(DataFormats.FileDrop);

            if (files == null) return;

            files.ToList().ForEach(s =>
            {
                if (viewModel.Workbooks.All(w => w.Path != s && FileTypeChecker.IsFileExtelType(s)))
                {
                    var newWB = new WorkbookWithItemsQnt(s);
                    newWB.WorksheetsForCountMaxRows = new List<string>(viewModel.WorksheetsToCopy);
                    viewModel.Workbooks.Add(newWB);
                }
            });

//            UpdateTotalItems();
            waitWindow.Close();
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
                    }
            }
            else if (e.Key == Key.Enter)
            {
                StartCombine();
            }
            
        }

        private void ButtonBase_OnClick(object sender, RoutedEventArgs e)
        {
            if (viewModel.Workbooks.Count == 0) return;
            if (viewModel.WorksheetsToCopy.Count == 0 && !viewModel.AllSheetsInOne) return;

            ExcelHelper.App.DisplayAlerts = false;

//            await CombineAsync();
            var waitWindow = new PleaseWaitWindow() { Owner = this };
            waitWindow.Show();
            waitWindow.Show();
            BlockUi();

            StartCombine();

            UnblockUi();
            ResetParams();
            waitWindow.Close();

            ExcelHelper.App.DisplayAlerts = true;
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
            ExcelApp.Visible = false;
            ExcelApp.EnableEvents = false;

            List<string> selectedWorksheets = null;

            if (viewModel.AllSheetsInOne)
                selectedWorksheets = viewModel.AllWorksheetsCollection.Distinct().ToList();
            else
                selectedWorksheets = viewModel.WorksheetsToCopy.Distinct().ToList();

            if (!selectedWorksheets.Any()) return;


            
            Workbook sampleWb = ExcelHelper.GetWorkbook(ExcelApp, viewModel.Workbooks.First().Path);
            Workbook resultWb;

            if (viewModel.AllSheetsInOne)
            {
                resultWb = ExcelHelper.CreateNewWorkbook(ExcelHelper.App);

                var resultWs = (Worksheet)resultWb.Worksheets[1];
                resultWs.Name = selectedWorksheets.First();

                var sourceWs =
                    sampleWb.Worksheets.Cast<Worksheet>().First();

                WriteWideHead(resultWs, sourceWs, viewModel.HeadSize);
            }
            else
            {
                resultWb = ExcelHelper.CreateNewWorkbook(ExcelHelper.App, (byte)selectedWorksheets.Count());
                //Create result worksheets
                for (int i = 1; i <= selectedWorksheets.Count(); i++)
                {
                    var resultWs = (Worksheet)resultWb.Worksheets[i];
                    resultWs.Name = selectedWorksheets[i - 1];

                    var sourceWs =
                        sampleWb.Worksheets.Cast<Worksheet>()
                            .FirstOrDefault(w => string.Equals(resultWs.Name, w.Name, StringComparison.OrdinalIgnoreCase));

                    WriteWideHead(resultWs, sourceWs, viewModel.HeadSize);
                }

            }
            
            var fillers =
                resultWb.Worksheets.Cast<Worksheet>().Select(w => new WorksheetFiller(w)).ToList();

            

            foreach (var workbookInfo in viewModel.Workbooks)
            {
                var wb = ExcelHelper.GetWorkbook(ExcelApp,workbookInfo.Path);

                foreach (var targetWs in selectedWorksheets)
                {
                    var sourceWs =
                        wb.Worksheets.Cast<Worksheet>()
                            .FirstOrDefault(
                                w => String.Equals(w.Name, targetWs, StringComparison.OrdinalIgnoreCase));

                    if (sourceWs == null) continue;

                    try
                    {
                        sourceWs.ShowAllData();
                    }
                    catch (Exception)
                    {
                        //ignored
                    }

                    WorksheetFiller filler;
                    if (viewModel.AllSheetsInOne)
                    {
                        filler = fillers.First();
                    }
                    else
                        filler = fillers.First(
                            f => string.Equals(f.WorksheetName, targetWs, StringComparison.OrdinalIgnoreCase));

                    filler.InsertOneToOneWorksheet(sourceWs, viewModel.HeadSize + 1);
                }
                wb.Close();
            }

            try
            {
                ExcelApp.EnableEvents = false;
                ExcelApp.Visible = true;
                resultWb.Activate();
                ((Worksheet)resultWb.Worksheets[1]).Activate();
            }
            catch (COMException)
            {
                
                return;
            }
            
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
            ExcelHelper.App = ExcelHelper.CreateNewApplication();
            viewModel = new ViewModel();
            DataContext = viewModel;
            WorkbooksListBox.ItemsSource = viewModel.Workbooks;
            SelectedWorksheetsListBox.ItemsSource = viewModel.WorksheetsToCopy;
            MyProgressBar.Value = 0;
        }

        private void ManualUpdateWindow()
        {
            UpdateTotalItems();
            AllWorksheetsListBox.ItemsSource = viewModel.AllWorksheetsCollection;
        }

        private void UpdateTotalItems()
        {
            Binding binding = new Binding();
            binding.ElementName = "TotalItemsQntTextBox";
            binding.Path = new PropertyPath("TotalItemsQuantity");
            TotalItemsQntTextBox.SetBinding(TextBox.TextProperty, binding);

//            TotalItemsQntTextBox.Text = viewModel.Workbooks.Sum(w => w.MaxRowsInWorkbook).ToString();
        }

        private void WriteWideHead(Worksheet targetWs, Worksheet soureWs, byte headSize)
        {
            for (int i = 1; i <= headSize; i++)
            {
                ((Range)targetWs.Rows[i]).EntireRow.Value2 = ((Range)soureWs.Rows[i]).EntireRow.Value2;
            }
        }
    }
}
