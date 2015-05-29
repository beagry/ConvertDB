using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Input;
using Converter;
using Converter.Template_workbooks;
using ExcelRLibrary;
using RwaySupportLibraly;

namespace UI
{
    /// <summary>
    /// Interaction logic for ConverterWindow.xaml
    /// </summary>
    public partial class ConverterWindow
    {
        private BooksToConvertViewModel viewModel;


        public ConverterWindow()
        {
            InitializeComponent();
            ResetWindow();
        }

        private void ListBox_Drop(object sender, DragEventArgs e)
        {
            var files = (string[])e.Data.GetData(DataFormats.FileDrop);

            if (files != null)
                files.ToList().ForEach(s =>
                {
                    if (viewModel.Workbooks.All(w => w.Path != s))
                    {
                        viewModel.Workbooks.Add(new SelectedWorkbook(s));
                    }
                });
        }

        private void Window_Closing(object sender, System.ComponentModel.CancelEventArgs e)
        {
            var window = sender as Window;
            if (window != null && window.Owner != null)
                window.Owner.Close();
        }

        private async void Button_Click(object sender, RoutedEventArgs e)
        {

            viewModel.EditMode = false;
            var wbAnalyzier = new WorkbooksAnalyzier(viewModel.WorkbooksType);

            viewModel.StartWork();
            await Task.Run(() =>
            {
                var paths = viewModel.Workbooks.Select(wb => wb.Path);
                wbAnalyzier.CheckWorkbooks(paths);
            });
            
            
            var worksheets = wbAnalyzier.WorksheetsInfos;
            var dict = wbAnalyzier.ComparedColumns;

            viewModel.EndWork();

            var w = new ColumnsCompareWindow(dict, worksheets);
            w.Closed += (o, args) => 
            {
                ResetWindow();
                this.Show();
                viewModel.EditMode = true;
            };

            w.Show();
            this.Hide();
        }

        private void ConverterWindow_OnKeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Delete)
            {
                var selItems = WorkbooksListBox.SelectedItems;
                if (selItems == null || selItems.Count == 0) return;

                foreach (var item in selItems.Cast<SelectedWorkbook>().ToList())
                {
                    viewModel.Workbooks.Remove(item);
                }
            }
        }

        private void ResetWindow()
        {
            viewModel = new BooksToConvertViewModel();

            DataContext = viewModel;

            foreach (Enum e in Enum.GetValues(typeof(XlTemplateWorkbookType)))
                WorkbookTypesComboBox.Items.Add(e.GetDescription());
        }
    }
}
