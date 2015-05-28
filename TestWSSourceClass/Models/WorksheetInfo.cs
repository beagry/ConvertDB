using System.Collections.Generic;
using System.Data;
using System.Linq;
using ExcelRLibrary;
using Microsoft.Office.Interop.Excel;
using DataTable = System.Data.DataTable;

namespace Converter.Models
{
    public class WorksheetInfo
    {
        public string Name { get; private set; }
        public ICollection<ColumnInfo> Columns  { get; private set; }
        public SelectedWorkbook Workbook { get; set; }


        public WorksheetInfo(DataTable dt):this()
        {
            var head = dt.Columns.Cast<DataColumn>().ToDictionary(k => dt.Columns.IndexOf(k) + 1, v => v.ColumnName);
            Name = dt.TableName;
            foreach (var pair in head)
                Columns.Add(new ColumnInfo(dt,pair.Key,pair.Value));
        }


        public WorksheetInfo(Worksheet ws):this()
        {
            var head = ws.ReadHead();

            foreach (var keyValuePair in head)
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