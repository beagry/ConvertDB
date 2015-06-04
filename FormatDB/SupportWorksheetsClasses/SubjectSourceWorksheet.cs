using System;
using System.Globalization;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using Microsoft.Office.Interop.Excel;
using Microsoft.Win32;
using DataRow = System.Data.DataRow;
using DataTable = System.Data.DataTable;

namespace Formater.SupportWorksheetsClasses
{
    public class SubjectSourceWorksheet
    {
        private readonly DataTable table;
        private const byte SourceNameColumnIndex = 1;
        private const byte LinkColumnIndex = 2;
        private const byte SubjectColumnIndex = 3;
        private const byte DefaultCityColumnIndex = 5;
        

        public SubjectSourceWorksheet(DataTable table)
        {
            this.table = table;
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

            //Создаём новый паттерн для поиска соответствия в нашей таблице
            pattern = Regex.Replace(pattern, @"\\d\{2\,3\}", digital.ToString(CultureInfo.InvariantCulture), RegexOptions.IgnoreCase);
            pattern = pattern.Replace(@"\\", @"\");
            reg = new Regex(pattern);

            var row = table.Rows.Cast<DataRow>().FirstOrDefault(r =>
            {
                var val = r[LinkColumnIndex - 1];
                if (val == null) return false;
                var m = reg.Match(val.ToString());
                return m.Success;
            });

            if (row == null) return string.Empty;

            var resCell = row[DefaultCityColumnIndex-1];
            var result = (resCell??"").ToString();
            return result;
        }

        public string GetSubjectBySourceLink(string sourceLink)
        {
            if (sourceLink == null) return string.Empty;
            var row = table.Rows.Cast<DataRow>().FirstOrDefault(r =>
            {
                var val = r[LinkColumnIndex - 1];
                if (val == null) return false;
                return (sourceLink.IndexOf(val.ToString(), StringComparison.OrdinalIgnoreCase) >= 0);

            });
            if (row == null) return string.Empty;

            var res = (row[SubjectColumnIndex-1]??"").ToString();
            return res;
        }
    }
}
