using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Excel;
using DataTable = System.Data.DataTable;
using Excel = Microsoft.Office.Interop.Excel;

namespace Converter
{
    static class ExcelExtensions
    {
        /// <summary>
        /// Метод возвращает текущий лист как DataTable
        /// </summary>
        /// <param name="worksheet">Лист для концертации</param>
        /// <returns></returns>
        public static DataTable GetDataTableFromWorksheet(this Worksheet worksheet)
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

                if (columnName == "NoName")
                    dataTable.Columns.Add(
                        new DataColumn());
                else
                    dataTable.Columns.Add(
                        new DataColumn(columnName));
            }


            //for each row
            for (var i = 2; i <= rangeArray.GetLength(0); i++)
            {
                DataRow dataRow = dataTable.NewRow();

                //fill row
                //for each column
                for (var j = 1; j <= rangeArray.GetLength(1); j++)
                {
                    dataRow[j-1] = rangeArray[i, j]??string.Empty;
                }
                dataTable.Rows.Add(dataRow);
            }
            return dataTable;
        }

        public static DataTable GetDataTableFromWorksheet(this Worksheet worksheet, int rows)
        {
            var datatable = worksheet.GetDataTableFromWorksheet();

            DataTable newTable = datatable.Clone();
            for (int i = 0; i <= rows; i++)
            {
                newTable.ImportRow(datatable.Rows[i]);
            }

            return newTable;
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

        [DllImport("user32.dll")]
        static extern int GetWindowThreadProcessId(int hWnd, out int lpdwProcessId);

        public static Process GetExcelProcess(Excel.Application excelApp)
        {
            int id;
            GetWindowThreadProcessId(excelApp.Hwnd, out id);
            return Process.GetProcessById(id);
        }
    }

    public class ExcelApp
    {
        public ExcelApp()
        {
            XlApplication = new Excel.Application();//GetExcelApplication();
            XlProcess = ExcelExtensions.GetExcelProcess(XlApplication);
        }

        /// <summary>
        /// COM объект текущего Excel приложения
        /// </summary>
        public Application XlApplication { get; private set; }

        /// <summary>
        /// Объект - процесс текущего Excel приложения
        /// </summary>
        public Process XlProcess { get; private set; }

        /// <summary>
        /// Метод возвращает книгу, открытую в текущем процессе
        /// </summary>
        /// <param name="workbookPath">путь к книге</param>
        /// <returns>COM объект workbook книги, открытой по переданному адресу в текущем процессе</returns>
        public Excel.Workbook OpenWorkbook(string workbookPath)
        {
            if (!File.Exists(workbookPath)) return null;
            

            XlProcess.StartInfo = new ProcessStartInfo(workbookPath);
            XlProcess.Start();

            var workbookName = Path.GetFileName(workbookPath);
            var workbook = XlApplication.Workbooks[workbookName];

            return workbook;
        }

        public static Microsoft.Office.Interop.Excel.Application GetExcelApplication()
        {

#if (DEBUG)
            return new Excel.Application { Visible = true, ScreenUpdating = true };
#endif
            Excel.Application xlApplication = null;
            try
            {
                xlApplication = (Excel.Application)Marshal.GetActiveObject("Excel.Application");
            }
            catch (COMException exception)
            {
                if (xlApplication == null)
                {
                    xlApplication = new Excel.Application() { Visible = true };
                }
                else
                {
                    throw;
                }
            }
            return xlApplication;
        }
    }

}