using System;
using System.Globalization;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using Microsoft.Office.Interop.Excel;

namespace Formater.SupportWorksheetsClasses
{
    public class SubjectSourceWorksheet
    {
        private Worksheet worksheet;
        private readonly Range nameRange;
        private Range subjectRange;
        private readonly Range linkRange;
        private long lastUsedRow;
        private const byte SourceNameColumnIndex = 1;
        private const byte LinkColumnIndex = 2;
        private const byte SubjectColumnIndex = 3;
        private const byte DefaultCityColumnIndex = 5;
        

        public SubjectSourceWorksheet(Worksheet worksheet)
        {

            this.worksheet = worksheet;// workbook.Worksheets.Cast<Worksheet>().FirstOrDefault(x => x.Name == WorksheetName);
            if (worksheet == null) return;
            try
            {
                worksheet.ShowAllData();
            }
            catch (COMException e)
            {
                if (e.HResult != -2146827284) throw;
            }
            var t5 = worksheet.UsedRange.Rows.Count;
            lastUsedRow = worksheet.Cells.SpecialCells(XlCellType.xlCellTypeLastCell).Row;

            nameRange =
                worksheet.Range[
                    worksheet.Cells[2, SourceNameColumnIndex], worksheet.Cells[lastUsedRow, SourceNameColumnIndex]];
            subjectRange =
                            worksheet.Range[
                                worksheet.Cells[2, SubjectColumnIndex], worksheet.Cells[lastUsedRow, SubjectColumnIndex]];
            linkRange =
                            worksheet.Range[
                                worksheet.Cells[2, LinkColumnIndex], worksheet.Cells[lastUsedRow, LinkColumnIndex]];


        }

        public string GetDefaultNearCityByLink(string sourceLink)
        {
            var pattern = @"(http://|^)(dom(\.)?)?(?<num>\d{2,3})\.ru";
            var reg = new Regex(pattern, RegexOptions.IgnoreCase);
            var match = reg.Match(sourceLink);

            if (!match.Success) return String.Empty;
            if (Regex.IsMatch(sourceLink, @"dom72.ru", RegexOptions.IgnoreCase) ||
                Regex.IsMatch(sourceLink, @"dom49.ru", RegexOptions.IgnoreCase)) return String.Empty;

            int digital;
            int.TryParse(match.Groups["num"].Value,out digital);
            
            //if (digital == 0) very bad

            //Создаём новый паттерн для поиска соответствия в нашей таблице
            pattern = Regex.Replace(pattern, @"\\d\{2\,3\}", digital.ToString(CultureInfo.InvariantCulture), RegexOptions.IgnoreCase);
            pattern = pattern.Replace(@"\\", @"\");
            reg = new Regex(pattern);

//            foreach (Range cell in linkRange.Cells.Cast<Range>())
//            {
//                if (cell.Value2 != null && reg.IsMatch())
//            }
            Range firstOrDefault = linkRange.Cells.Cast<Range>()
                .FirstOrDefault(
                    cell =>
                        cell.Value2 != null &&
                        reg.IsMatch(cell.Value2.ToString()));
                        //sourceLink.IndexOf(cell.Value2.ToString(), StringComparison.OrdinalIgnoreCase) >= 0);

            if (firstOrDefault == null) return String.Empty;

            var resCell = (Range) worksheet.Cells[firstOrDefault.Row, DefaultCityColumnIndex];
            var result = resCell.Value2 == null ? String.Empty : resCell.Value2.ToString();
            return result;
        }

        [Obsolete("Субъект лучше доставать через метод GetSubjectBySourceLink", true)]
        public string GetSubjectBySourceName(string sourceName)
        {
            if (sourceName == null) return String.Empty;
            var firstOrDefault =
                nameRange.Cells.Cast<Range>()
                    .Where(x => x.Value2 != null)
                    .FirstOrDefault(
                        x => sourceName.IndexOf(x.Value2.ToString(), StringComparison.OrdinalIgnoreCase) >= 0);
            if (firstOrDefault == null) return String.Empty;

            var res = ((Range)worksheet.Cells[firstOrDefault.Row, SubjectColumnIndex]).Value2;
            return res;
        }

        public string GetSubjectBySourceLink(string sourceLink)
        {
            if (sourceLink == null) return string.Empty;
            var firstOrDefault = linkRange.Cells.Cast<Range>().Where(x => x.Value2 != null).FirstOrDefault(x => sourceLink.IndexOf(x.Value2.ToString(), StringComparison.OrdinalIgnoreCase) >= 0);
            if (firstOrDefault == null) return string.Empty;

            var cellRow = firstOrDefault.Row;

            var res = ((Range)worksheet.Cells[cellRow, SubjectColumnIndex]).Value2;
            return res;
        }

        public void CloseWorkbook()
        {
            Workbook workbook = worksheet.Parent;
            workbook.Close(false);
        }
    }
}
