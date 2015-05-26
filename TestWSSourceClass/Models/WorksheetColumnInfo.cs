using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using Microsoft.Office.Interop.Excel;

namespace Converter.Models
{
    public class ColumnInfo
    {
        public static byte ExamplesQnt = 10;
        public int Index { get; private set; }
        public String Name { get; private set; }
        public ICollection<string> ValuesExamples { get; set; }


        public ColumnInfo(Worksheet ws, int index, string name):this(index,name)
        {
            Index = index;
            Name = name;
            SetValuesExamples(ws);
        }

        public ColumnInfo(int index, string name)
        {
            Index = index;
            Name = name;
        }

        public ColumnInfo()
        {

        }


        private void SetValuesExamples(Worksheet ws)
        {
            var columnRange = (Range) ws.Columns[Index, Type.Missing];

            ValuesExamples =
                columnRange.Cells.Cast<Range>().Select(c => c.Value2)
                    .Skip(1)
                    .Take(500)
                    .Where(v => v != null && v.ToString() != "")
                    .Select(v => v.ToString())
                    .Distinct()
                    .Take(ExamplesQnt)
                    .Cast<string>().ToList();
        }
    }
}