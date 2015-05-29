using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using ExcelRLibrary;
using Microsoft.Office.Interop.Excel;

namespace Converter.Models
{
    internal class WorksheetToCopy
    {
        private readonly long lastUsedRow;
        private readonly Worksheet worksheet;
        private byte headRow;

        public WorksheetToCopy(Worksheet worksheet, byte headRow = 1, bool oneToOne = false)
        {
            this.worksheet = worksheet;
            lastUsedRow = worksheet.GetLastUsedRow();
            this.headRow = headRow;

            HeadsDictionary = oneToOne
                ? worksheet.ReadHead(headRow).ToDictionary(k => k.Key, v => v.Key.ToString())
                : worksheet.ReadHead(headRow);
        }

        public Dictionary<int, string> HeadsDictionary { get; private set; }
        public byte FirstRowWithData { get; set; }

        public long DataRowsQnt
        {
            get
            {
                if (FirstRowWithData == 0) return 0;
                return lastUsedRow - FirstRowWithData + 1;
            }
        }

        public void CopyColumn(int column, Range firstTargetCell, bool withFormat = false)
        {
            var copyRange =
                (Range)worksheet.Range[worksheet.Cells[FirstRowWithData, column], worksheet.Cells[lastUsedRow, column]];
            var pasteRange = GetRangeProjection(copyRange, firstTargetCell);

            if (withFormat)
            {
                var groupedByFormat =
                    copyRange.Cast<Range>()
                        .Take(100)
                        .GroupBy(c => c.NumberFormat)
                        .ToDictionary(g => g.Key, v => v.Count());
                if (groupedByFormat.Count > 1)
                {
                    if (groupedByFormat.ContainsKey("General"))
                        groupedByFormat["General"] -= 3;
                    if (groupedByFormat.ContainsKey("@"))
                        groupedByFormat["@"] -= 2;
                }

                var format = groupedByFormat.First(v => v.Value == groupedByFormat.Values.Max()).Key;
                pasteRange.NumberFormat = format;
            }

            if (copyRange.Cells.Cast<Range>().Count() == 1)
            {
                if (copyRange.Value2 != null)
                {
                    if (pasteRange.Value2 == null)
                        pasteRange.Value2 = copyRange.Value2;
                    else
                        pasteRange.Value2 += copyRange.Value2;
                }
                return;
            }

            object[,] copyArray = copyRange.Value2;
            object[,] pasteArray = pasteRange.Value2;

            for (var i = 1; i <= copyArray.GetLength(0); i++)
            {
                if (pasteArray[i, 1] == null)
                    pasteArray[i, 1] = copyArray[i, 1];
                else
                    pasteArray[i, 1] += ", " + copyArray[i, 1];
            }

            try
            {
                pasteRange.Value2 = pasteArray;
            }
            catch (COMException e)
            {
                if (e.HResult == -2146827284)
                {
                    if (pasteArray != null)
                    {
                        var pattern = "^=";
                        var reg = new Regex(pattern);

                        //Исправляем формат "=аывав" на "аывав"
                        for (var i = 1; i < copyArray.GetLength(0); i++)
                        {
                            if (copyArray[i, 1] == null) continue;
                            var newVal = copyArray[i, 1].ToString();
                            if (!reg.IsMatch(newVal)) continue;
                            newVal = reg.Replace(newVal, "");

                            pasteArray[i, 1] = newVal;
                        }

                        try
                        {
                            pasteRange.Value2 = pasteArray;
                        }
                        catch (COMException)
                        {
                            //ignored
                        }
                    }
                }
                else
                    throw;
            }
        }

        private Range GetRangeProjection(Range range, Range firstCell)
        {
            if (range.Columns.Count > 1) return null;

            var rowsQnt = range.Cells.Count;
            var lastCell = firstCell.Offset[rowsQnt - 1, range.Columns.Count - 1];

            var projectionWS = (Worksheet) firstCell.Parent;
            var projectionRange = projectionWS.Range[firstCell, lastCell];

            return projectionRange;
        }
    }
}