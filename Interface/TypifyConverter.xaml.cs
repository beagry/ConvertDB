using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Input;
using Converter;
using Converter.Template_workbooks;
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
            viewModel = new BooksToConvertViewModel();

            DataContext = viewModel;

            foreach (Enum e in Enum.GetValues(typeof (XlTemplateWorkbookTypes)))
                WorkbookTypesComboBox.Items.Add(e.GetDescription());
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


        private void Button_Click(object sender, RoutedEventArgs e)
        {
            var dict = new Dictionary<string, List<string>>();
            List<WorksheetInfo> unbindedColumns = new List<WorksheetInfo>();


            var w = new ColumnsCompareWindow(dict, unbindedColumns);
            w.Closed += (o, args) => this.Show();
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
    }
}
