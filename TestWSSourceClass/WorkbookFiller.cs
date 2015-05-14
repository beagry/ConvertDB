using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using Converter.Models;
using ExcelRLibrary;
using Microsoft.Office.Interop.Excel;

namespace Converter
{
    /// <summary>
    /// Простой помощни для записи информации в книгу.
    /// </summary>
    public class WorksheetFiller
    {
        private long lastUsedRow;
        private int lastUsedColumn;
        private bool oneToOneMode;

        private readonly Worksheet fillingWorksheet;
        private Dictionary<int, string> headsDictionary;


        public Dictionary<string, List<string>> RulesDictionary { get; set; }
        public string WorksheetName { get { return fillingWorksheet.Name; } }




        public WorksheetFiller(Worksheet fillingWorksheet, Dictionary<string, List<string>> rulesDictionary):this(fillingWorksheet)
        {
            RulesDictionary = rulesDictionary;
            headsDictionary = fillingWorksheet.ReadHead();
        }


        public WorksheetFiller(Worksheet fillingWorksheet)
        {
            this.fillingWorksheet = fillingWorksheet;
            lastUsedRow = fillingWorksheet.GetLastUsedRow();
            lastUsedColumn = fillingWorksheet.GetLastUsedColumnByRow();
        }


        public void InsertOneToOneWorksheet(Worksheet ws, int firstRowWithData = 1)
        {

            var copyRange = ws.Range[ws.Cells[firstRowWithData, 1], ws.Cells[ws.GetLastUsedRow(), ws.GetLastUsedColumn()]];
            var cellTopaste = fillingWorksheet.Cells[lastUsedRow,1];
            copyRange.Copy(cellTopaste);
            lastUsedRow += copyRange.Rows.Count + 1;
//            DeleteLastEmptyRows();

            Marshal.FinalReleaseComObject(copyRange);
            Marshal.FinalReleaseComObject(cellTopaste);

        }

        public void InsertWorksheet(Worksheet ws, int firstRowWithData = 1, bool copyFormat = false)
        {
            CheckRulesDict();

            var wsWithData = new WorksheetToCopy(ws) { FirstRowWithData = (byte) firstRowWithData };

            //Каждую колонку из копируемого листа
            foreach (var indexNamePair in wsWithData.HeadsDictionary)
            {
                //ищем подготовленную для неё колонку вставки
                int indexToPaste = oneToOneMode? indexNamePair.Key : GetColumnIndexToPaste(indexNamePair.Value);

                //если правил нет, колонку вставляем в конец книги
                if (indexToPaste == 0)
                {
                    indexToPaste = lastUsedColumn++;
                    headsDictionary.Add(indexToPaste, indexNamePair.Value);
                }

                var cellToPaste = fillingWorksheet.Cells[lastUsedRow + 1, indexToPaste] as Range;
                var copyColumnIndex = indexNamePair.Key;

                wsWithData.CopyColumn(copyColumnIndex,cellToPaste,copyFormat);
            }

            lastUsedRow = fillingWorksheet.GetLastUsedRow();
            DeleteLastEmptyRows();
        }

        private void CheckRulesDict()
        {
            if (RulesDictionary == null)
                SetOneToOneRulesDict();
        }

        private void SetOneToOneRulesDict()
        {
            headsDictionary = fillingWorksheet.ReadHead().ToDictionary(k => k.Key, v => v.Key.ToString());
            RulesDictionary = headsDictionary.ToDictionary(k => k.Key.ToString(), v => new List<string> { v.Key.ToString() }); 
        }

        private void DeleteLastEmptyRows()
        {
            while (
                ((Range)fillingWorksheet.Range[
                    fillingWorksheet.Cells[lastUsedRow, 1],
                    fillingWorksheet.Cells[lastUsedRow, fillingWorksheet.Cells.SpecialCells(XlCellType.xlCellTypeLastCell).Column]]).Cells
                    .Cast<Range>().All(cl => cl.Value2 == null))
            {
                lastUsedRow --;
            }
        }

        private int GetColumnIndexToPaste(string columnNameToSearch)
        {
            if (!RulesDictionary.Any(
                kv => kv.Value.Any(s => string.Equals(s, columnNameToSearch, StringComparison.OrdinalIgnoreCase))))
                return 0;

            var columnNameToPaste = RulesDictionary.
                First(
                    kv =>
                        kv.Value.Any(s => string.Equals(s, columnNameToSearch, StringComparison.OrdinalIgnoreCase)))
                .Key;

            return
                headsDictionary.First(
                    kv => string.Equals(kv.Value, columnNameToPaste, StringComparison.OrdinalIgnoreCase)).Key;
        }
    }
}
