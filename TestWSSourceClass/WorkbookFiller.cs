using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using Converter.Models;
using ExcelRLibrary;
using Microsoft.Office.Interop.Excel;
using OfficeOpenXml;
using DataTable = System.Data.DataTable;

namespace Converter
{

    public interface IFiller
    {
        Dictionary<string,List<string>> RulesDictionary { get; set; }
    }

    interface IWorksheetFiller:IFiller
    {
        void InsertOneToOneWorksheet(Worksheet ws, int firstRowWithData = 1);
        void InsertWorksheet(Worksheet ws, int firstRowWithData = 1, bool copyFormat = false);
    }

    interface IEPPlusWorksheetFiller:IFiller
    {
        ExcelWorksheet Worksheet{ get; }
        void AppendDataTable(DataTable dt);
    }

    /// <summary>
    /// Простой помощни для записи информации в книгу.
    /// </summary>
    public class WorksheetFiller:IWorksheetFiller,IEPPlusWorksheetFiller
    {
        private long lastUsedRow;
        private int lastUsedColumn;
        private bool oneToOneMode;
        private int colNum = 1;
        private const string colName = "AddColumn_";

        private readonly Worksheet fillingWorksheet;
        private Dictionary<int, string> headsDictionary;


        public Dictionary<string, List<string>> RulesDictionary { get; set; }
        public string WorksheetName { get { return fillingWorksheet.Name; } }


        public WorksheetFiller(ExcelWorksheet worksheet, Dictionary<string, List<string>> rulesDictionary):this()
        {
            this.Worksheet = worksheet;
            headsDictionary = worksheet.ReadHead();
            RulesDictionary = rulesDictionary;
            lastUsedColumn = Worksheet.Dimension.Columns;
            lastUsedRow = Worksheet.Dimension.Rows;
        }

        public WorksheetFiller(Worksheet fillingWorksheet, Dictionary<string, List<string>> rulesDictionary):this(fillingWorksheet)
        {
            RulesDictionary = rulesDictionary;
            headsDictionary = fillingWorksheet.ReadHead();
        }


        public WorksheetFiller(Worksheet fillingWorksheet):this()
        {
            this.fillingWorksheet = fillingWorksheet;
            lastUsedRow = fillingWorksheet.GetLastUsedRow();
            lastUsedColumn = fillingWorksheet.GetLastUsedColumnByRow();
        }

        private WorksheetFiller()
        {
            RulesDictionary = new Dictionary<string, List<string>>();
            headsDictionary = new Dictionary<int, string>();
        }


        public void InsertOneToOneWorksheet(Worksheet ws, int firstRowWithData = 1)
        {

            var copyRange = ws.Range[ws.Cells[firstRowWithData, 1], ws.Cells[ws.GetLastUsedRow(), ws.GetLastUsedColumn()]];
            var cellTopaste = fillingWorksheet.Cells[lastUsedRow,1];
            copyRange.Copy(cellTopaste);
            lastUsedRow += copyRange.Rows.Count + 1;

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
            //проверяем наличие искомой колонки в списке с правилами вставка
            if (!RulesDictionary.Any(
                kv => kv.Value.Any(s => string.Equals(s, columnNameToSearch, StringComparison.OrdinalIgnoreCase))))
                return 0;

            //Ну и извлекаем название колонки, в которую переносится переданная колонка
            var columnNameToPaste = RulesDictionary.
                First(
                    kv =>
                        kv.Value.Any(s => string.Equals(s, columnNameToSearch, StringComparison.OrdinalIgnoreCase)))
                .Key;

            return
                headsDictionary.First(
                    kv => string.Equals(kv.Value, columnNameToPaste, StringComparison.OrdinalIgnoreCase)).Key;
        }



        public ExcelWorksheet Worksheet { get; private set; }

        public void AppendDataTable(DataTable dt)
        {
            var tableColumns = dt.ReadHead();
            
            //Поколоночно
            foreach (var indexNamePair in tableColumns)
            {

                var pasteColumnName = indexNamePair.Value;

                //ищем подготовленную для неё колонку вставки
                var indexToPaste = GetColumnIndexToPaste(indexNamePair.Value);

                //если правил нет, колонку вставляем в конец книги
                if (indexToPaste == 0)
                {
                    indexToPaste = ++lastUsedColumn;

                    if (!headsDictionary.ContainsValue(pasteColumnName))
                        headsDictionary.Add(indexToPaste, pasteColumnName);
                    else
                    {
                        Worksheet.Cells[1, indexToPaste].Value = colName + colNum++;
                        headsDictionary.Add(indexToPaste, indexNamePair.Value);
                    }
                }

                
                var copyColumn = dt.Columns[indexNamePair.Key - 1];

                //Вставляем всю колонку построчно
                for (var i = 0; i < dt.Rows.Count; i++)
                {
                    var row = dt.Rows[i];
                    var cellToPaste = Worksheet.Cells[(int)lastUsedRow + 1 + i, indexToPaste];
                    cellToPaste.Value = row[copyColumn];
                }
            }

            lastUsedRow += dt.Rows.Count;
        }
    }
}
