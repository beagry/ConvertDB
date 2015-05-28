using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using Microsoft.Office.Interop.Excel;
using DataTable = System.Data.DataTable;

namespace Converter.Models
{
    public class ColumnInfo
    {
        public static byte ExamplesQnt = 10;

        public ColumnInfo(Worksheet ws, int index, string name) : this(index, name)
        {
            Index = index;
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

        public ColumnInfo(DataTable table, int index, string name)
            : this(index, name)
        {
            Index = index;
            SetValuesExamples(table);
        }

        public int Index { get; private set; }
        public string Name { get; private set; }
        public ICollection<string> ValuesExamples { get; set; }

        private void SetValuesExamples(DataTable table)
        {
            var column = table.Columns[Index - 1];
            ValuesExamples = table.Rows.Cast<DataRow>().Take(500)
                .Select(r => r[column])
                .Where(o => o != null)
                .Select(s => s.ToString())
                .Distinct().Take(ExamplesQnt).ToList();
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