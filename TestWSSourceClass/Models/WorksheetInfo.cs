using System.Collections.Generic;
using ExcelRLibrary;
using Microsoft.Office.Interop.Excel;

namespace Converter.Models
{
    public class WorksheetInfo
    {
        public string Name { get; private set; }
        public ICollection<ColumnInfo> Columns  { get; private set; }
        public SelectedWorkbook Workbook { get; set; }


        public WorksheetInfo(Worksheet ws)
        {
            var head = ws.ReadHead();

            Columns = new List<ColumnInfo>();
            foreach (KeyValuePair<int, string> keyValuePair in head)
                Columns.Add(new ColumnInfo(ws, keyValuePair.Key, keyValuePair.Value));
        }
        public WorksheetInfo(Dictionary<int,string> head)
        {
            Columns = new List<ColumnInfo>();
            foreach (KeyValuePair<int, string> keyValuePair in head)
                Columns.Add(new ColumnInfo(keyValuePair.Key, keyValuePair.Value));
        }
        public WorksheetInfo(string name, ICollection<ColumnInfo> columns, SelectedWorkbook workbook)
        {
            Name = name;
            Columns = columns;
            Workbook = workbook;
        }
        public WorksheetInfo(string name, ICollection<ColumnInfo> columns)
        {
            Name = name;
            Columns = columns;
            Workbook = null;
        }
        public WorksheetInfo()
        {
            Columns = new List<ColumnInfo>();
            Workbook = null;
        }
    }
}