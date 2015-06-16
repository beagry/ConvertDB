using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Collections.Specialized;
using System.Data;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Input;
using Converter;
using ExcelRLibrary;
using OfficeOpenXml;
using Telerik.Windows.Controls;
using Action = System.Action;
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

        public MainWindow()
        {
            InitializeComponent();

            ResetParams();
            DataContext = viewModel;
        }

        private void ListBox_Drop(object sender, DragEventArgs e)
        {
            var files = (string[])e.Data.GetData(DataFormats.FileDrop);

            if (files == null) return;

            viewModel.StartWork();

            files.ToList().ForEach(s =>
            {
                if (!viewModel.Workbooks.All(w => w.Path != s && FileTypeChecker.IsFileExtelType(s))) return;

                var newWB = new WorkbookWithItemsQnt(s)
                {
                    WorksheetsForCountMaxRows = new List<string>(viewModel.WorksheetsToCopy)
                };
                viewModel.Workbooks.Add(newWB);
            });

            viewModel.EndWork();
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

        private async void ButtonBase_OnClick(object sender, RoutedEventArgs e)
        {
            if (viewModel.Workbooks.Count == 0) return;
            if (viewModel.WorksheetsToCopy.Count == 0) return;

            viewModel.StartWork();

            await CombineAsync();
            ResetParams();

            viewModel.EndWork();

        }

        private async Task CombineAsync()
        {
            await Task.Run((Action) StartCombine);

        }

        private void StartCombine()
        {
            var combiner = new WorkbookCombiner()
            {
                WorksheetsToCombine = viewModel.WorksheetsToCopy.Distinct().ToList(),
                WorkbooksPaths = viewModel.Workbooks.Select(w => w.Path).ToList()
            };
            var pkcg = combiner.Combine();
            pkcg.SaveWithDialog();
        }


        private void ResetParams()
        {
            ExcelHelper.App = ExcelHelper.CreateNewApplication();
            viewModel = new ViewModel();
            DataContext = viewModel;
            WorkbooksListBox.ItemsSource = viewModel.Workbooks;
            SelectedWorksheetsListBox.ItemsSource = viewModel.WorksheetsToCopy;
        }
    }

    class WorkbookCombiner
    {
        public List<string> WorksheetsToCombine { get; set; }
        public List<string> WorkbooksPaths { get; set; }

        public int HeadSize { get; set; }

        public WorkbookCombiner()
        {
            WorksheetsToCombine = new List<string>();
            WorkbooksPaths = new List<string>();
            HeadSize = 1;
        }

        public ExcelPackage Combine()
        {
            if (!WorksheetsToCombine.Any()) return null;
            if (!WorkbooksPaths.Any()) return null;

            var pckg = new ExcelPackage();

            var reader = new ExcelReader();
            var sampleWb = reader.ReadExcelFile(WorkbooksPaths.First());

            List<ExcelWorksheet> worksheets = new List<ExcelWorksheet>();

            //Create Result WB
            foreach (
                var sourceTable in
                    WorksheetsToCombine.Select(s => sampleWb.Tables.Cast<DataTable>().First(t => t.TableName.Equals(s))))
            {
                var ws = pckg.Workbook.Worksheets.Add(sourceTable.TableName);

                var head = sourceTable.ReadHead();
                ws.WriteHead(head);
                worksheets.Add(ws);
            }

            var fillers = worksheets.Select(w => new WorksheetFiller(w)).ToList();

            foreach (var wbDataSet in WorkbooksPaths.Select(s => reader.ReadExcelFile(s)))
            {
                foreach (var wsName in WorksheetsToCombine)
                {
                    var sourceTable =
                        wbDataSet.Tables.Cast<DataTable>()
                            .FirstOrDefault(t => t.TableName.Equals(wsName, StringComparison.OrdinalIgnoreCase));

                    if (sourceTable == null) continue;


                    var filler = fillers.First(
                        f => string.Equals(f.WorksheetName, wsName, StringComparison.OrdinalIgnoreCase));

                    filler.InsertOneToOneWorksheet(sourceTable);
                }
            }
            return pckg;
        }

    }
}
