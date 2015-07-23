using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using Converter.Template_workbooks;
using Converter.Template_workbooks.EFModels;
using ExcelRLibrary.TemplateWorkbooks;
using TemplateWorkbook = Converter.Template_workbooks.EFModels.TemplateWorkbook;

namespace Converter
{
    public class SourceWs
    {
/*
        private const int TakeFirstItemsQuantity = 300;
*/
        private readonly List<int> checkedColumnsList;

        /// <summary>
        ///     Key = номер столбца, который будет скопирован, Value = Название колонки Куда будет скопирован столбец
        /// </summary>
        private readonly Dictionary<int, string> columnsDictionary = new Dictionary<int, string>();
        private readonly Dictionary<int, string> head;
        private readonly TemplateWorkbook wb;
        private readonly DataTable wsTable;

        /// <summary>
        ///     Самый продуктивный конструктор
        /// </summary>
        /// <param name="table"></param>
        /// <param name="templateWorkbook"></param>
        public SourceWs(DataTable table, TemplateWorkbook templateWorkbook)
            : this()
        {
            wb = templateWorkbook;
            wsTable = table;
            head = wsTable.Columns.Cast<DataColumn>()
                .ToDictionary(k => wsTable.Columns.IndexOf(k) + 1, v => v.ColumnName);
        }

        private SourceWs()
        {
            checkedColumnsList = new List<int>();
        }

        public Dictionary<string, List<string>> ResultDictionary
        {
            get
            {
                return columnsDictionary
                    .Select(kp => new {ColumnCopy = head.First(hk => hk.Key == kp.Key).Value, ColumnPaste = kp.Value})
                    .GroupBy(obj => obj.ColumnPaste, o => o.ColumnCopy)
                    .ToDictionary(k => k.Key, v => v.ToList());
            }
        }

        public void CheckColumns()
        {
            //
            //Общие колонки
            //

            FindOneToOneColumn();
            GetBindedColumnsFromDb();
            TryToFindTemplateColumnsFromDbData();
        }

        private void GetBindedColumnsFromDb()
        {
            var columns = wb.Columns;

            var tableColumns =
                wsTable.Columns.Cast<DataColumn>()
                    .Select(c => new {Index = wsTable.Columns.IndexOf(c) + 1, Name = c.ColumnName}).ToList();
            foreach (var column in columns.Where(c => c.BindedColumns.Any()))
            {
                for (var i = tableColumns.Count - 1; i >= 0; i--)
                {
                    var tableColumn = tableColumns[i];
                    if (!column.BindedColumns.Any(bc => bc.Name.Equals(tableColumn.Name))) continue;

                    columnsDictionary.Add(tableColumn.Index, column.CodeName);
                    tableColumns.Remove(tableColumn);
                }
            }
        }

        /// <summary>
        ///     Поиск колонок с точно такими же именами как в конечной книге
        /// </summary>
        private void FindOneToOneColumn()
        {
            var columns = wb.Columns;

            foreach (var column in columns)
            {
                if (!GetColumnNumberByColumnName(column.CodeName, new List<string> { column.CodeName },true))
                    GetColumnNumberByColumnName(column.CodeName, new List<string> { column.Name },true);
            }
        }

        private void TryToFindTemplateColumnsFromDbData()
        {
            var columns = wb.Columns;

            foreach (var column in columns)
            {
                var maskList = column.SearchCritetias.Select(s => s.Text).ToList();
                var columnCode = column.CodeName;
                GetColumnNumberByColumnName(columnCode, maskList);
            }
        }

        /// <summary>
        ///     Метод находит колонку с полным или частичным совпадением в имени
        /// </summary>
        /// <param name="columnCode">Название колонки для записи результата</param>
        /// <param name="masks">Маски для сопоставления</param>
        /// <param name="fullSimilar">Обязательно полное совпадение</param>
        /// <returns></returns>
        private bool GetColumnNumberByColumnName(string columnCode, IReadOnlyCollection<string> masks, bool fullSimilar = false)
        {
            if (masks.Count == 0) return false;
            
            JustColumn cl = null;
            var tableColumns =
                wsTable.Columns.Cast<DataColumn>()
                .Select(col => new JustColumn {Index = col.Ordinal + 1, CodeName = col.ColumnName})
                .ToList();

            //Поиск колонки с ПОЛНЫМ совпалением по одному из критериев маски поиска
            foreach (var mask in masks)
            {
                cl = tableColumns.Where(x => !columnsDictionary.ContainsKey(x.Index))
                    .FirstOrDefault(x => string.Equals(x.CodeName, mask, StringComparison.CurrentCultureIgnoreCase));

                if (cl != null) break;
            }

            //Поиск колонки с ЧАСТИЧНЫМ совпалением по одному из критериев маски поиска
            if (cl == null && !fullSimilar)
            {
                foreach (var mask in masks)
                {
                    cl = tableColumns.Where(x => !columnsDictionary.ContainsKey(x.Index))
                        .FirstOrDefault(x => x.CodeName.IndexOf(mask, StringComparison.OrdinalIgnoreCase) > -1);

                    if (cl != null) break;
                }
            }

            if (cl == null) return false;

            checkedColumnsList.Add(cl.Index);
            columnsDictionary.Add(cl.Index, columnCode);
            return true;
        }

    }
}

// ReSharper restore SuggestUseVarKeywordEvident