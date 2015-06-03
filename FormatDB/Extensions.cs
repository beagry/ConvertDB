using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Runtime.InteropServices;
using Formater.SupportWorksheetsClasses;
using Microsoft.Office.Interop.Excel;
using DataTable = System.Data.DataTable;

namespace Formater
{
    static class ExcelExtensions
    {
        /// <summary>
        /// Метод возвращает текущий лист как DataTable.
        /// Индекс первой строки и колонки = 0.
        /// </summary>
        /// <param name="worksheet">Лист для концертации</param>
        /// <returns></returns>
        public static DataTable ToDataTable(this Worksheet worksheet)
        {
            //To Reset UsedRange
            var noth = worksheet.UsedRange.Rows.Count;
            noth = worksheet.UsedRange.Columns.Count;

            object[,] rangeArray = worksheet.Range[worksheet.Cells[1,1],worksheet.Cells.SpecialCells(XlCellType.xlCellTypeLastCell)].Value2;
            DataTable dataTable = new DataTable();

            //create heads
            for (var j = 1; j <= rangeArray.GetLength(1); j++)
            {
                string columnName = (string) (rangeArray[1, j] ?? "NoName");

                dataTable.Columns.Add(
                    new DataColumn(columnName));

                
            }

            
            //for each row
            for (var i = 2; i <= rangeArray.GetLength(0); i++)
            {
                DataRow dataRow = dataTable.NewRow();

                //fill each cell
                for (var j = 1; j <= rangeArray.GetLength(1); j++)
                {
                    dataRow[j-1] = rangeArray[i, j]??string.Empty;
                }
                dataTable.Rows.Add(dataRow);
            }
            return dataTable;
        }

        public static void DeleteEmptyRows(this _Worksheet worksheet)
        {
            var lasRow = worksheet.Cells.SpecialCells(XlCellType.xlCellTypeLastCell).Row;

            var s = worksheet.UsedRange.Rows.Count; //reset usedrange
            for (var row = lasRow; row >= 0; row--)
            {
                Range rowRange =
                    worksheet.UsedRange.Range[
                        worksheet.Cells[row, 1],
                        worksheet.Cells[row, worksheet.UsedRange.Cells.SpecialCells(XlCellType.xlCellTypeLastCell).Column]];

                var j = rowRange.Cells.Cast<Range>().Count(cellInRow => cellInRow.Value2 != null);
                if (j != 0) continue; // if row has not empty cell we move next
                rowRange.EntireRow.Delete(XlDeleteShiftDirection.xlShiftUp); //empty row to delete
            }
        }

        //Появилось само!! Что это может быть?!
        public static IEnumerable<TSource> DistinctBy<TSource, TKey>
            (this IEnumerable<TSource> source, Func<TSource, TKey> keySelector)
        {
            HashSet<TKey> seenKeys = new HashSet<TKey>();
            foreach (TSource element in source)
            {
                if (seenKeys.Add(keySelector(element)))
                {
                    yield return element;
                }
            }
        }

        /// <summary>
        /// Метод возвращает таблицу, с той же структурой, но с содержанием из выборки
        /// </summary>
        /// <returns></returns>
        public static DataTable GetCustomDataTable(this DataTable table, Func<DataRow, bool> pred)
        {
            if (table == null) return null;

            var resTable = table.Clone();
            foreach (var row in table.Rows.Cast<DataRow>())
            {
                if (pred(row))
                resTable.ImportRow(row);
            }
            return resTable;
        }

        internal enum CellColors
        {
            BadColor,
            GoodColor,
            Clear
        }

        public static void ColorCell(this Range cell,CellColors cellColor)
        {
            switch (cellColor)
            {
                case CellColors.BadColor:
                    cell.Interior.Color = Color.Crimson;
                    break;
                case CellColors.GoodColor:
                    cell.Interior.Color = Color.Aquamarine;
                    break;
                case CellColors.Clear:
                    cell.Interior.ColorIndex = 0;
                    break;

            }
        }

        public static void ReleaseComObject(Object obj)
        {
            Marshal.ReleaseComObject(obj);
            GC.Collect();
        }

    }

    /// <summary>
    /// Методы расширения для DatTable
    /// </summary>
    static class DataTableExtensions
    {
        public static DataTable GetCustomDataTable(this DataTable dataTable, params SearchParams[] searchParams)
        {
            //var result = new List<string>();
            var newTable = dataTable.Copy();


            //Сортировка для поиска от общего к частному
            var newSearchParams = searchParams.OrderBy(x => x.SearchColumn);

            foreach (SearchParams @params in newSearchParams)
            {
                //Ищем все строки, в которых в ячейках по искомому столбцу строки содержат искомое значение
                var searchColumn = OKTMOWorksheet.GetExcelColumn(@params.SearchColumn) - 1;
                var searchString = @params.SearchString;
                newTable =
                    newTable.GetCustomDataTable( //Метод создания новой таблицы по условию
                        row => row[searchColumn].ToString().IndexOf(searchString, StringComparison.CurrentCultureIgnoreCase) >= 0);
            }
            Console.WriteLine(newTable.Rows.Count);
            //Из полученной таблицы достаём нужную нам колонку
            return newTable;
        }
    }

    static class OtherExtentions
    {
        public static bool Contains(this string source, string toCheck, StringComparison comp)
        {
            return source.IndexOf(toCheck, comp) >= 0;
        }
    }
    
}