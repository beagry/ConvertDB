using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelRLibrary;
using Microsoft.Office.Interop.Excel;

namespace Converter
{
    class WorksheetsColumnsAnalyzer
    {
        
        //Возвращает список уникальных столбцов от всех полученных книг
        //В списке так же присутствуют образцы информации
    }

    public class WorksheetInfo
    {
        public string Name { get; private set; }
        public ICollection<WorksheetColumnInfo> Columns  { get; private set; }



        public WorksheetInfo(Worksheet ws)
        {
            var head = WorksheetExtentions.GetHeadsDictionary(ws);

            Columns = new List<WorksheetColumnInfo>();
            foreach (KeyValuePair<int, string> keyValuePair in head)
                Columns.Add(new WorksheetColumnInfo(ws, keyValuePair.Key, keyValuePair.Value));
        }
        public WorksheetInfo(Dictionary<int,string> head)
        {
            Columns = new List<WorksheetColumnInfo>();
            foreach (KeyValuePair<int, string> keyValuePair in head)
                Columns.Add(new WorksheetColumnInfo(keyValuePair.Key, keyValuePair.Value));
        }
        public WorksheetInfo(string name, ICollection<WorksheetColumnInfo> columns)
        {
            Name = name;
            Columns = columns;
        }
        public WorksheetInfo()
        {
            
        }
    }

    public class WorksheetColumnInfo
    {
        public static byte ExamplesQnt = 10;
        public int Index { get; private set; }
        public String Name { get; private set; }
        public IEnumerable<string> ValuesExamples { get; set; }


        public WorksheetColumnInfo(Worksheet ws, int index, string name):this(index,name)
        {
            Index = index;
            Name = name;
            SetValuesExamples(ws);
        }
        public WorksheetColumnInfo(int index, string name)
        {
            Index = index;
            Name = name;
        }
        public WorksheetColumnInfo()
        {
        }


        private void SetValuesExamples(Worksheet ws)
        {
            ValuesExamples = new List<string>();
            var i = 1;
            var columnRange = ws.Columns[Index, Type.Missing] as Range;

            Debug.Assert(columnRange != null, "columnRange != null");
            ValuesExamples =
                // ReSharper disable once PossibleNullReferenceException
                (IEnumerable<string>) columnRange.Cells.Cast<Range>()
                    .Where(c => c.Value2 != null && c.Value2.ToString() != "")
                    .Take(ExamplesQnt)
                    .Select(c => c.Value2.ToString()).AsEnumerable();
        }
    }
}
