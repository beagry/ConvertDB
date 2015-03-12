using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text.RegularExpressions;
using Converter.Template_workbooks;
using ExcelRLibrary;
using Microsoft.Office.Interop.Excel;

namespace Converter
{
    public class WorkbookTypifier<T> where T : TemplateWorkbook, new ()
    {
        public Dictionary<string,IEnumerable<string>> RulesDictionary { get; set; }

        public WorkbookTypifier()
        {
            RulesDictionary = null;
            
        }
        public WorkbookTypifier(Dictionary<string, IEnumerable<string>> columnsDictionary)
        {
            RulesDictionary = columnsDictionary;
        }

        public Workbook CombineToSingleWorkbook(IEnumerable<string> workbooksPaths)
        {
            var helper = new ExcelHelper();
            //создать пустую книгу
            var newWb = helper.CreateNewWorkbook();
            var ws = newWb.Worksheets[1] as Worksheet;
            var wsWriter = new WorksheetFiller(ws,RulesDictionary);

            //* написовать шапку из шаблона
            var templateHead = new T().TemplateColumns.ToDictionary(k => k.Index, v => v.CodeName);
            WorksheetExtentions.WriteHead(ws,templateHead);
            
            //* поочередно открывать книги из списка
            foreach (var openWs in helper.GetWorkbooks(workbooksPaths).Select(wb => wb.Worksheets[1]).Cast<Worksheet>())
            {
                wsWriter.InsertWorksheet(openWs);
            }

            return newWb;
        }        
    }

    class WorksheetFiller
    {
        private long lastUsedRow;
        private int lastUsedColumn;
        private Worksheet worksheet;
        private Dictionary<int, string> headsDictionary;
        private readonly Dictionary<string, IEnumerable<string>> rulesDictionary; 

        public WorksheetFiller(Worksheet worksheet, Dictionary<string,IEnumerable<string>> rulesDictionary )
        {
            this.worksheet = worksheet;
            this.rulesDictionary = rulesDictionary;

            lastUsedRow = worksheet.GetLastUsedRow();
            lastUsedColumn = worksheet.GetLastUsedColumn();
            headsDictionary = WorksheetExtentions.GetHeadsDictionary(worksheet);
        }

        public void InsertWorksheet(Worksheet ws)
        {
            var copyWs = new WorksheetToCopy(ws);

            foreach (var indexNamePair in copyWs.HeadsDictionary)
            {
                //используя правила из словаря
                var indexToPaste = GetColumnIndexToPaste(indexNamePair.Value);

                //Paste to the end
                if (indexToPaste == 0)
                {
                    indexToPaste = lastUsedColumn++;
                    headsDictionary.Add(indexToPaste, indexNamePair.Value);
                }

                var firstCellToPaste = worksheet.Cells[lastUsedRow + 1, indexToPaste] as Range;
                var copyColumnIndex = indexNamePair.Key;
                copyWs.CopyColumn(copyColumnIndex,firstCellToPaste);
                //* копировать колонки из открытых книг в созданную книгу
            }
        }

        /// <summary>
        /// Возвращает номер столбца, в который будет осуществляться вставка
        /// </summary>
        /// <param name="columnNameToSearch">название колонки, для которой нужно найти место</param>
        /// <returns></returns>
        private int GetColumnIndexToPaste(string columnNameToSearch)
        {
            if (!rulesDictionary.Any(
                kv => kv.Value.Any(s => string.Equals(s, columnNameToSearch, StringComparison.OrdinalIgnoreCase))))
                return 0;

            var columnNameToPaste = rulesDictionary.
                First(
                    kv =>
                        kv.Value.Any(s => string.Equals(s, columnNameToSearch, StringComparison.OrdinalIgnoreCase)))
                .Key;

            return
                headsDictionary.First(
                    kv => string.Equals(kv.Value, columnNameToPaste, StringComparison.OrdinalIgnoreCase)).Key;
        }
    }

    class WorksheetToCopy
    {
        private Worksheet worksheet;
        private long lastUsedRow ;
        private byte headRow;
        public Dictionary<int, string> HeadsDictionary { get; private set; }
//        public Dictionary<int,string> HeadsDictionary { get; private set; }

        public WorksheetToCopy(Worksheet worksheet, byte headRow = 1)
        {
            this.worksheet = worksheet;
            HeadsDictionary = WorksheetExtentions.GetHeadsDictionary(worksheet,headRow);
            this.headRow = headRow;
            lastUsedRow = worksheet.GetLastUsedRow();
        }

        public void CopyColumn(int column, Range firstTargetCell)
        {
            var copyRange = worksheet.Range[worksheet.Cells[headRow+1,column], worksheet.Cells[lastUsedRow,column]] as Range;
            object[,] copyArray = copyRange.Value2;
            var pasteRange = GetRangeProjection(copyRange, firstTargetCell);
            object[,] pasteArray = pasteRange.Value2;

            for (int i = 1; i < copyArray.GetLength(0); i++)
            {
                if (pasteArray[i,1] == null)
                    pasteArray[i, 1] = copyArray[i, 1];
                else
                    pasteArray[i, 1] += ", " + copyArray[i, 1];
            }

            try
            {
                pasteRange.Value2 = pasteArray;
            }
            catch (Exception e)
            {
                if (e.HResult == -2146827284)
                {
                    if (pasteArray != null)
                    {
                        var pattern = "^=";
                        var reg = new Regex(pattern);

                        //Исправляем формат "=аывав" на "аывав"
                        for (int i = 1; i < copyArray.GetLength(0); i++)
                        {
                            if (copyArray[i, 1] == null) continue;
                            var newVal = copyArray[i, 1].ToString();
                            if (!reg.IsMatch(newVal)) continue;
                                newVal = reg.Replace(newVal, "");

                            pasteArray[i, 1] = newVal;
                        }

                        pasteRange.Value2 = pasteArray;
                    }
                }
                else
                    throw;
            }

        }

        private Range GetRangeProjection(Range range, Range firstCell)
        {
            if (range.Columns.Count > 1 ) return null;

            var rowsQnt = range.Cells.Count;

            var projectionWS = range.Parent as Worksheet;
            Debug.Assert(projectionWS != null, "projectionWS != null");
            // ReSharper disable once PossibleNullReferenceException
            var projectionRange = projectionWS.Range[firstCell, firstCell.Offset[rowsQnt - 1, 0]] as Range;

            return projectionRange;

        }
    }
}
