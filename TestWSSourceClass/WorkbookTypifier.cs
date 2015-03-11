using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Media.Animation;
using ExcelRLibrary;
using Microsoft.Office.Interop.Excel;
using DataTable = System.Data.DataTable;

namespace Converter
{
    /*
     
     */
    public class WorkbookTypifier<T> where T : Template_workbooks.TemplateWorkbook, new ()
    {
        private readonly Dictionary<int, string> _templateHead;
        public Dictionary<string,IEnumerable<string>> RulesDictionary { get; set; }

        public WorkbookTypifier()
        {
            RulesDictionary = null;
            _templateHead = new T().TemplateColumns.ToDictionary(k => k.Index, v => v.CodeName);
        }
        public WorkbookTypifier(Dictionary<string, IEnumerable<string>> columnsDictionary)
        {
            RulesDictionary = columnsDictionary;
            _templateHead = new T().TemplateColumns.ToDictionary(k => k.Index, v => v.CodeName);
        }

        public Workbook CombineToSingleWorkbook(IEnumerable<string> workbooksPaths)
        {
            var helper = new ExcelHelper();
            //создать пустую книгу
            var newWb = helper.CreateNewWorkbook();
            var ws = newWb.Worksheets[1] as Worksheet;

            //* написовать шапку из шаблона
            WorksheetExtentions.WriteHead(ws,_templateHead);
            
            //* поочередно открывать книги из списка
            foreach (Workbook workbook in helper.GetWorkbooks(workbooksPaths))
            {
                //Refresh usedrange
                // ReSharper disable once PossibleNullReferenceException
                Console.WriteLine(ws.UsedRange.Rows.Count);
                var lastUsedRow = ws.Cells.SpecialCells(XlCellType.xlCellTypeLastCell).Row;
                var lastUserColumn = ws.Cells.SpecialCells(XlCellType.xlCellTypeLastCell).Column;

                var openWs = workbook.Worksheets[1] as Worksheet;
                var openWbHeads = WorksheetExtentions.GetHeadsDictionary(openWs);

                foreach (var indexNamePair in openWbHeads)
                {
                    var indexToPaste = GetColumnIndexToPaste(indexNamePair.Value);

                    //Paste to then end
                    if (indexToPaste == 0)
                        indexToPaste = lastUserColumn++;


                }

                //* копировать колонки из открытых книг в созданную книгу

                //* используя правила из словаря


            }

            return newWb;
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
                _templateHead.First(
                    kv => string.Equals(kv.Value, columnNameToPaste, StringComparison.OrdinalIgnoreCase)).Key;
        }
    }
}
