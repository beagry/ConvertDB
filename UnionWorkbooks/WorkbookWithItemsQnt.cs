using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Windows;
using Converter.Annotations;
using ExcelRLibrary;

namespace UnionWorkbooks
{
    sealed class WorkbookWithItemsQnt:SelectedWorkbook, INotifyPropertyChanged
    {
        private Application app;
        private Dictionary<string, long> worksheetsRowsQntDictionary;
        private List<string> worksheetsForCountMaxRows;
        private string type;


        public long MaxRowsInWorkbook
        {
            get
            {
                var list =
                    worksheetsRowsQntDictionary.Where(kv => WorksheetsForCountMaxRows.Any(s => s.Equals(kv.Key)))
                        .ToList();

                return list.Count==0?0:list.Max(kv => kv.Value);
            }
        }

        public List<string> WorksheetsForCountMaxRows
        {
            get { return worksheetsForCountMaxRows; }
            set
            {
                worksheetsForCountMaxRows = value;
                OnPropertyChanged("MaxRowsInWorkbook");
            }
        }

        public List<string> WorksheetsNamesList { get; set; }


        public WorkbookWithItemsQnt(string path):base(path)
        {
            worksheetsRowsQntDictionary = new Dictionary<string, long>();
            WorksheetsForCountMaxRows = new List<string>();
            Init();
        }

        private void Init()
        {
            var wb = ExcelHelper.GetWorkbook(ExcelHelper.App, Path);

            var wsList = wb.Worksheets.Cast<Microsoft.Office.Interop.Excel.Worksheet>();

            worksheetsRowsQntDictionary =  wsList.ToDictionary(toKey => toKey.Name,w =>
            {
                Console.WriteLine(w.UsedRange.Rows.Count);
                try
                {
                    return (long) w.Cells.SpecialCells(Microsoft.Office.Interop.Excel.XlCellType.xlCellTypeLastCell).Row - 1;
                }
                catch (Exception)
                {

                    return 0;
                }

            });

            WorksheetsNamesList = wb.Worksheets.Cast<Microsoft.Office.Interop.Excel.Worksheet>().Select(w => w.Name).ToList();
            wb.Close();
        }


        public event PropertyChangedEventHandler PropertyChanged;

        [NotifyPropertyChangedInvocator]
        private void OnPropertyChanged([CallerMemberName] string propertyName = null)
        {
            var handler = PropertyChanged;
            if (handler != null) handler(this, new PropertyChangedEventArgs(propertyName));
        }
    }
}