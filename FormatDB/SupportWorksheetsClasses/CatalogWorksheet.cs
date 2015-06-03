using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;

namespace Formater.SupportWorksheetsClasses
{
    class CatalogWorksheet
    {
        private Worksheet worksheet;
        private const byte CodeColumnIndex = 2;
        private const byte NameColumnIndex = 3;
        private const byte ContentColumnIndex = 5;

        private static int lastUsedRow;

        private Range codeColumnRange;
        private Range nameColumnRange;
        private Range contentColumnRange;

        public CatalogWorksheet(Worksheet ws)
        {
            worksheet = ws;
            try
            {
                worksheet.ShowAllData();
            }
            catch (COMException e)
            {
                if (e.HResult != -2146827284) throw;
            }                
            //worksheet.Cells.EntireRow.Hidden = false;
            //if (worksheet.EnableAutoFilter) worksheet.EnableAutoFilter = false;
            //worksheet.Cells.EntireRow.AutoFill();
            
            //For reset UsedRange
            var t2 = worksheet.UsedRange.Rows.Count;

            try
            {
                lastUsedRow = worksheet.Cells.SpecialCells(XlCellType.xlCellTypeLastCell).Row;
            }
            catch (COMException)
            {

                MessageBox.Show("Выйдите в Экселе из режима редактирования.","Операция прервана",MessageBoxButtons.OK,MessageBoxIcon.Exclamation);
                Environment.Exit(0);
            }
            
            codeColumnRange =
                worksheet.Range[worksheet.Cells[2, CodeColumnIndex], worksheet.Cells[lastUsedRow, CodeColumnIndex]];
            nameColumnRange =
                worksheet.Range[worksheet.Cells[2, NameColumnIndex], worksheet.Cells[lastUsedRow, NameColumnIndex]];
            contentColumnRange =
                worksheet.Range[worksheet.Cells[2, ContentColumnIndex], worksheet.Cells[lastUsedRow, ContentColumnIndex]];
        }

        public List<string> GetContentByCode(string code)
        {
            //Если ввели неверный код
            if (codeColumnRange.Cast<Range>().FirstOrDefault(x => x.Value2 == code) == null) return null;

            var resultList = codeColumnRange.Cast<Range>()
                .Where(x => x.Value2 == code)
                .Select(x => x.Offset[0, ContentColumnIndex - CodeColumnIndex].Value2.ToString())
                .Cast<String>().ToList();

            return resultList;
        }

        public List<String> GetContentByName(string name)
        {
            //Если ввели неверный код
            if (nameColumnRange.Cast<Range>().
                FirstOrDefault(x => x.Value2 == name) == null) 
                return null;

            var resultList = nameColumnRange.Cast<Range>()
                .Where(x => x.Value2 == name)
                .Select(x => x.Offset[0, ContentColumnIndex - NameColumnIndex].Value2.ToString())
                .Cast<String>().ToList();

            return resultList;
        }

        public void CloseWorkbook()
        {
            Workbook workbook = worksheet.Parent;
            workbook.Close(false);
        }
    }
}