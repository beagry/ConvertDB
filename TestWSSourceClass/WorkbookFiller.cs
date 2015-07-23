using System;
using System.Collections.Generic;
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
        Dictionary<string, List<string>> RulesDictionary { get; set; }
    }

    internal interface IEPPlusWorksheetFiller : IFiller
    {
        ExcelWorksheet Worksheet { get; }
        void AppendDataTable(DataTable dt);
    }

    /// <summary>
    ///     Простой помощни для записи информации в книгу.
    /// </summary>
    public class WorksheetFiller : IEPPlusWorksheetFiller
    {
        private const string colName = "AddColumn_";
        private int colNum = 1;
        private Dictionary<int, string> headsDictionary;
        private int lastUsedColumn;
        private int lastUsedRow;

        

        public WorksheetFiller(ExcelWorksheet worksheet, Dictionary<string, List<string>> rulesDictionary) : this(worksheet)
        {
            RulesDictionary = rulesDictionary;
        }

        public WorksheetFiller(ExcelWorksheet worksheet) : this()
        {
            Worksheet = worksheet;
            headsDictionary = worksheet.ReadHead();
            lastUsedColumn = Worksheet.Dimension.End.Column;
            lastUsedRow = Worksheet.Dimension.End.Row;
        }

        private WorksheetFiller()
        {
            RulesDictionary = new Dictionary<string, List<string>>();
        }




        public string WorksheetName
        {
            get { return Worksheet.Name; }
        }
        public ExcelWorksheet Worksheet { get; private set; }
        public Dictionary<string, List<string>> RulesDictionary { get; set; }


        public void AppendDataTable(DataTable dt)
        {
            var tableColumns = dt.ReadHead();

            //Поколоночно
            foreach (var indexNamePair in tableColumns)
            {
                var pasteColumnName = indexNamePair.Value;

                //ищем подготовленную для неё колонку вставки
                var indexToPaste = GetColumnIndexToPaste(pasteColumnName);

                //если правил нет, колонку вставляем в конец книги
                if (indexToPaste == 0)
                {
                    indexToPaste = ++lastUsedColumn;
                    headsDictionary.Add(indexToPaste, pasteColumnName);
                    Worksheet.Cells[1, indexToPaste].Value = pasteColumnName;
                    RulesDictionary.Add(pasteColumnName, new List<string> {pasteColumnName});
                }

                var copyColumn = dt.Columns[indexNamePair.Key - 1];

                //Вставляем всю колонку построчно
                for (var i = 0; i < dt.Rows.Count; i++)
                {
                    var row = dt.Rows[i];
                    var cellToPaste = Worksheet.Cells[(int) lastUsedRow + 1 + i, indexToPaste];
                    var currVal = (cellToPaste.Value ?? "").ToString();
                    if (currVal != "")
                    {
                        if (row[copyColumn].ToString() != "")
                            cellToPaste.Value = currVal + " | " + row[copyColumn];
                    }
                    else
                        cellToPaste.Value = row[copyColumn];
                }
            }

            lastUsedRow += dt.Rows.Count;
        }


        public void InsertOneToOneWorksheet(DataTable sourceTable)
        {
            var cellTopaste = Worksheet.Cells[++lastUsedRow, 1];
            cellTopaste.LoadFromDataTable(sourceTable, false);

            lastUsedRow += sourceTable.Rows.Count -  1;
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

            if (!headsDictionary.ContainsValue(columnNameToPaste))
            {
                var indexToPaste = ++lastUsedColumn;
                headsDictionary.Add(indexToPaste, columnNameToPaste);
                Worksheet.Cells[1, indexToPaste].Value = columnNameToPaste;
                return indexToPaste;
            }

            return
                headsDictionary.First(
                    kv => string.Equals(kv.Value, columnNameToPaste, StringComparison.OrdinalIgnoreCase)).Key;
        }
    }
}