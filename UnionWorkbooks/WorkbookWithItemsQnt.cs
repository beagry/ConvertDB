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
        private Dictionary<string, int> worksheetsRowsQntDictionary;
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
            worksheetsRowsQntDictionary = new Dictionary<string, int>();
            WorksheetsForCountMaxRows = new List<string>();
            Init();
        }

        private void Init()
        {
            var reader = new ExcelReader();
            var wbDataSet = reader.ReadExcelFile(Path);


            worksheetsRowsQntDictionary =
                wbDataSet.Tables.Cast<System.Data.DataTable>().ToDictionary(toKey => toKey.TableName, w =>
                    w.Rows.Count);

            WorksheetsNamesList = wbDataSet.Tables.Cast<System.Data.DataTable>().Select(dt => dt.TableName).ToList(); 
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