using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.Linq;
using System.Runtime.CompilerServices;
using Converter.Properties;
using ExcelRLibrary;
using Microsoft.Office.Interop.Excel;

namespace UnionWorkbooks
{
    sealed class WorkbookWithItemsQnt:SelectedWorkbook, INotifyPropertyChanged
    {
        private Dictionary<string, long> worksheetsRowsQntDictionary;
        private List<string> worksheetsForCountMaxRows;


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
//            Workbook wb = null;
            var wb = ExcelHelper.GetWorkbook(ExcelHelper.App, Path);
            if (wb == null)
            {
                var newApp = ExcelHelper.GetApplication();
                wb = newApp.Workbooks[System.IO.Path.GetFileName(Path)];
                ExcelHelper.App = newApp;
            }

            Debug.Assert(wb != null);

            List<Worksheet> wsList = null;
            if (System.IO.Path.GetExtension(Path) == ".csv")
                wsList = new List<Worksheet>(){wb.Worksheets[1]};
            else
                wsList = wb.Worksheets.Cast<Worksheet>().ToList();

            worksheetsRowsQntDictionary =  wsList.ToDictionary(toKey => toKey.Name,w =>
            {
                Console.WriteLine(w.UsedRange.Rows.Count);
                try
                {
                    return (long) w.Cells.SpecialCells(XlCellType.xlCellTypeLastCell).Row - 1;
                }
                catch (Exception)
                {

                    return 0;
                }

            });

            WorksheetsNamesList = wb.Worksheets.Cast<Worksheet>().Select(w => w.Name).ToList();
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