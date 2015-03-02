using System.Linq;
using Microsoft.Office.Interop.Excel;

namespace Converter
{
    //.EntireRow.Delete(XlDeleteShiftDirection.xlShiftUp);
    public static class  WorksheetExtensions
    {
        public static void DeleteEmptyRows(this _Worksheet worksheet)
        {
            int lasRow = worksheet.Cells.SpecialCells(XlCellType.xlCellTypeLastCell).Row;
            foreach (Range cell in worksheet.Range[worksheet.Cells[2,1],worksheet.Cells[lasRow,1]])
            {
                Range row =
                    worksheet.UsedRange.Range[
                        worksheet.Cells[cell.Row, 1],
                        worksheet.Cells[cell.Row, worksheet.UsedRange.Cells.SpecialCells(XlCellType.xlCellTypeLastCell).Column]];
                int j = row.Cast<Range>().Count(cellInRow => cellInRow.Value2 != null);
                if (j !=  0) continue;
                row.EntireRow.Delete(XlDeleteShiftDirection.xlShiftUp);
            }
        }
//        public static IEnumerable<TSource> DistinctBy<TSource, TKey>
//            (this IEnumerable<TSource> source, Func<TSource, TKey> keySelector)
//        {
//            HashSet<TKey> seenKeys = new HashSet<TKey>();
//            foreach (TSource element in source)
//            {
//                if (seenKeys.Add(keySelector(element)))
//                {
//                    yield return element;
//                }
//            }
//        }
    }
}
