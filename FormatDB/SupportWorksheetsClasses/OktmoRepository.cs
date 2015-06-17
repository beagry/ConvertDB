using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.Linq;
using System.Linq.Expressions;
using System.Text.RegularExpressions;
using ExcelRLibrary;
using ExcelRLibrary.SupportEntities.Oktmo;
using PatternsLib;

namespace Formater.SupportWorksheetsClasses
{
    public enum OKTMOColumn
    {
        //Название колонок с иерархией от общего к частному
        Subject = 1,
        Region = 2,
        Settlement = 3,
        NearCity = 4,
        TypeOfNearCity = 5
    }

    public interface IOktmoReposiroty
    {
        IEnumerable<OktmoRow> GetSubjectRows(string subj);
        IEnumerable<OktmoRow> GetRowsByRegion(string reg);
        IEnumerable<OktmoRow> GetRowsBySettlement(string sett);
        IEnumerable<OktmoRow> GetRowsByNearCity(string city);

        IEnumerable<OktmoRow> GetRowsBySearchParams(OktmoRow searchParamsRow); 


        /// <summary>
        ///     Возвращает региональный ценр переданного субъекта
        /// </summary>
        /// <param name="regionFullName">Название субъекта</param>
        /// <returns></returns>
        string GetDefaultRegCenter(string regionFullName);

        /// <summary>
        ///     Возвращает региональный ценрт, проверенный и приведенный к ОКТМО
        /// </summary>
        /// <param name="regionFullName"></param>
        /// <returns></returns>
        string GetDefaultRegCenterFullName(string regionFullName, ref string cityName);
    }

    public class OKTMORepository : IOktmoReposiroty
    {
        private static readonly Dictionary<OKTMOColumn, byte> classificatorColumnDictionary = new Dictionary
            <OKTMOColumn, byte>
        {
            {OKTMOColumn.Subject, OKTMOColumnsFilter.Default.Subject},
            {OKTMOColumn.Region, OKTMOColumnsFilter.Default.Region},
            {OKTMOColumn.Settlement, OKTMOColumnsFilter.Default.Settlement},
            {OKTMOColumn.NearCity, OKTMOColumnsFilter.Default.NearCity},
            {OKTMOColumn.TypeOfNearCity, OKTMOColumnsFilter.Default.TypeOfNearCity}
        };

        private readonly DataTable regCTable;
        private readonly Dictionary<string, DataTable> reserveDataTables = new Dictionary<string, DataTable>();
        private readonly DataTable table;
        private readonly List<OktmoRow> oktmoRows; 

        public OKTMORepository(DataSet ds, string mainWsName)
        {
            table = ds.Tables.Cast<DataTable>().First(t => t.TableName.Equals(mainWsName));
            regCTable = ds.Tables.Cast<DataTable>().FirstOrDefault(t => t.TableName.EqualNoCase("РегЦентры"));
            oktmoRows = table.Rows.Cast<DataRow>().Select(r => new OktmoRow()
            {
                Subject = (r[GetExcelColumn(OKTMOColumn.Subject)-1]??"").ToString(),
                Region =(r[GetExcelColumn(OKTMOColumn.Region)-1]??"").ToString(),
                Settlement = (r[GetExcelColumn(OKTMOColumn.Settlement)-1]??"").ToString(),
                NearCity = (r[GetExcelColumn(OKTMOColumn.NearCity)-1]??"").ToString(),
                TypeOfNearCity = (r[GetExcelColumn(OKTMOColumn.TypeOfNearCity) - 1] ?? "").ToString(),
            }).ToList();
        }

        public static ColumnNumbers Columns
        {
            get
            {
                var columns = new ColumnNumbers
                {
                    Subject = 3,
                    Region = 4,
                    Settlement = 5,
                    NearCity = 7,
                    TypeOfNearCity = 8
                };

                return columns;
            }
        }

        public DataTable Table
        {
            get { return table; }
        }

        
        public string GetDefaultRegCenter(string regionFullName)
        {
            if (string.IsNullOrEmpty(regionFullName)) return string.Empty;
            if (regCTable == null) return string.Empty;
            const int regionTableColumn = 0;
            const int regCenterTableColumn = 1;

            var row =
                regCTable.Rows.Cast<DataRow>()
                    .FirstOrDefault(
                        r =>
                            string.Equals(r[regionTableColumn].ToString(), regionFullName,
                                StringComparison.OrdinalIgnoreCase));
            if (row == null)
                return string.Empty;

            var res = row[regCenterTableColumn].ToString();
            res = res.Replace("город", "");
            res = res.Trim();

            return res;
        }

        public string GetDefaultRegCenterFullName(string regionFullName, ref string cityName)
        {
            return String.Empty;
            var regCenterName = GetDefaultRegCenter(regionFullName);
            if (string.IsNullOrEmpty(regCenterName)) return string.Empty;

            var regCentSpec = new ExpressionSpecification<OktmoRow>(row => string.Equals(row.NearCity,regCenterName,StringComparison.OrdinalIgnoreCase) );
            var cityRegCentSpec = new ExpressionSpecification<OktmoRow>(row => string.Equals(row.TypeOfNearCity,"город",StringComparison.OrdinalIgnoreCase));
            var spec = regCentSpec.And(cityRegCentSpec);

            var rows = oktmoRows.Where(r => spec.IsSatisfiedBy(r)).DistinctBy(r => r.Region).ToList();
            if (rows.Count() != 1) return  String.Empty;
            var singleRow = rows.Single();

            var fullName = singleRow != null? singleRow.Region: String.Empty;

            cityName = regCenterName;
            return fullName;
        }

        public bool StringMatchInColumn(DataTable table, string s, OKTMOColumn column)
        {
            if (string.IsNullOrEmpty(s)) return false;
            if (table == null)
                table = this.table;

            var res =
                table.Rows.Cast<DataRow>()
                    .Any(
                        row => row[GetExcelColumn(column) - 1] is string &&
                               string.Equals(ExcelLocationRow.ReplaceYo(row[GetExcelColumn(column) - 1].ToString()), s,
                                   StringComparison.OrdinalIgnoreCase));
            return res;
        }

        public IEnumerable<OktmoRow> GetSubjectRows(string subj)
        {
            var spec = new SubjectSpecification(subj);
            return oktmoRows.Where(oktmoRow => spec.IsSatisfiedBy(oktmoRow));
        }

        public IEnumerable<OktmoRow> GetRowsByRegion(string reg)
        {
            var spec = new ExpressionSpecification<OktmoRow>(row => string.Equals(row.Region,reg,StringComparison.OrdinalIgnoreCase));
            return oktmoRows.Where(oktmoRow => spec.IsSatisfiedBy(oktmoRow));
        }

        public IEnumerable<OktmoRow> GetRowsBySettlement(string sett)
        {
            var spec = new ExpressionSpecification<OktmoRow>(row => string.Equals(row.Settlement, sett, StringComparison.OrdinalIgnoreCase));
            return oktmoRows.Where(oktmoRow => spec.IsSatisfiedBy(oktmoRow));
        }

        public IEnumerable<OktmoRow> GetRowsByNearCity(string city)
        {
            var spec = new ExpressionSpecification<OktmoRow>(row => string.Equals(row.NearCity, city, StringComparison.OrdinalIgnoreCase));
            return oktmoRows.Where(oktmoRow => spec.IsSatisfiedBy(oktmoRow));
        }

        public IEnumerable<OktmoRow> GetRowsBySearchParams(OktmoRow searchParamsRow)
        {
            var spec = new OktmoSpecifications(searchParamsRow);
            return oktmoRows.Where(oktmoRow => spec.IsSatisfiedBy(oktmoRow));
        }

        /// <summary>
        ///     Аналог VlookUp in Excel
        /// </summary>
        /// <param name="searchParams"></param>
        /// <returns></returns>
        public DataTable GetCustomDataTable(params SearchParams[] searchParams)
        {
            var dataTable = table.Copy();


            //Сортировка для поиска от общего к частному
            foreach (var @params in searchParams)
            {
                //Ищем все строки, в которых в ячейках по искомому столбцу строки содержат искомое значение
                var searchColumn = GetExcelColumn(@params.SearchColumn) - 1;
                var searchString = @params.SearchString;
                dataTable =
                    dataTable.GetCustomDataTable( //Метод создания новой таблицы по условию
                        row =>
                            string.Equals(row[searchColumn].ToString(), searchString,
                                StringComparison.CurrentCultureIgnoreCase)); //Полное совпадение
            }
            Debug.Assert(dataTable.Rows.Count > 0);
            //Из полученной таблицы достаём нужную нам колонку
            return dataTable;
        }

        public DataTable GetCustomDataTable(DataTable currenTable, params SearchParams[] searchParams)
        {
            if (string.IsNullOrEmpty(searchParams[0].SearchString)) return currenTable;
            //var result = new List<string>();
            DataTable dataTable = null;
            if (currenTable == null) currenTable = table;

            Debug.Assert(searchParams.Count() == 1); //Пока работаем только с одним критерием поиска

            //Сортировка для поиска от общего к частному
            //searchParams = searchParams.OrderBy(x => x.SearchColumn);

            foreach (var @params in searchParams)
            {
                //Ищем все строки, в которых в ячейках по искомому столбцу строки содержат искомое значение
                var searchColumn = GetExcelColumn(@params.SearchColumn) - 1;
                var searchString = @params.SearchString;

                //Try to use reserved tables
                if (@params.SearchColumn.Equals(OKTMOColumn.Subject) &&
                    reserveDataTables.ContainsKey(searchString)) //Если в поиске субъект
                    dataTable = reserveDataTables[searchString];

                //Если в поиске муниципальное образование, то по одному образованию может быть несколько субъектом
                //Смотрим что найденная таблица содержит тот же субъект что и ?
                else if (@params.SearchColumn.Equals(OKTMOColumn.Region) &&
                         reserveDataTables.ContainsKey(searchString))
                {
                    var column = searchColumn;
                    //Сравниваем зарезервированную таблицу с ижу имеющейся таблицей
                    //Важно чтобы субъект по искомому муниОбразованию совпадал
                    var reservedSubjects =
                        reserveDataTables[searchString].Rows.Cast<DataRow>()
                            .Select(row => row[GetExcelColumn(OKTMOColumn.Subject) - 1].ToString())
                            .Distinct()
                            .ToList();
                    var currentSubject =
                        currenTable.Rows.Cast<DataRow>()
                            .Where(row => row[column].ToString() == searchString)
                            .Select(row => row[GetExcelColumn(OKTMOColumn.Subject) - 1].ToString())
                            .Distinct()
                            .ToList();
                    if (reservedSubjects.Count == 1 && currentSubject.Count == 1 &&
                        reservedSubjects[0] == currentSubject[0])
                        dataTable = reserveDataTables[searchString];
                }

                if (dataTable == null)
                {
                    dataTable =
                        currenTable.GetCustomDataTable(
                            row =>
                                //Только полное совпадление макси поиска с значением ячейки
                                string.Equals(ExcelLocationRow.ReplaceYo(row[searchColumn].ToString()), searchString,
                                    StringComparison.CurrentCultureIgnoreCase));
                    if ((@params.SearchColumn.Equals(OKTMOColumn.Subject) ||
                         @params.SearchColumn.Equals(OKTMOColumn.Region)) &&
                        !reserveDataTables.ContainsKey(searchString))
                        reserveDataTables.Add(searchString, dataTable);
                }
            }

            return dataTable;
        }

        private static byte GetExcelColumn(Enum searchColumn)
        {
            if (!classificatorColumnDictionary.ContainsKey((OKTMOColumn) searchColumn)) return 0;

            return classificatorColumnDictionary[(OKTMOColumn) searchColumn];
        }

        public List<string> GetContentByValue(OKTMOColumn contentColumn, string searchString, OKTMOColumn searchColumn,
            string searchString2, OKTMOColumn searchColumn2)
        {
            var result = new List<string>();

            //dont forget that datatable.columns[0] == worksheet.columns[1]
            foreach (var row in table.Rows.Cast<DataRow>())
            {
                if (
                    row[classificatorColumnDictionary[searchColumn] - 1].ToString()
                        .IndexOf(searchString, StringComparison.CurrentCultureIgnoreCase) >= 0 &&
                    row[classificatorColumnDictionary[searchColumn2] - 1].ToString()
                        .IndexOf(searchString2, StringComparison.CurrentCultureIgnoreCase) >= 0)
                {
                    result.Add((string) row[(int) contentColumn - 1]);
                }
            }

            return result.Distinct().ToList();
        }

        /// <summary>
        ///     Возвращает значение ячейки из выбранного столбца, с дополненным окончанием. Игнорирует регистр
        /// </summary>
        /// <param name="searchString">Искомое имя</param>
        /// <param name="searchColumn">Колонка поиска</param>
        /// <param name="type">Тип искомого наименования. Используется преимущественно для Муниципального Образования</param>
        /// <returns></returns>
        public string GetFullName(string searchString, OKTMOColumn searchColumn, string type = "")
        {
            var pattern = searchString.Trim() + "(\\b|$)";


            var results = oktmoRows.Select(r => GetPropValueByName(r, searchColumn))
                        .Where(val => Regex.IsMatch(val, pattern, RegexOptions.IgnoreCase)).Distinct().ToList();


            if (results.Count == 0) return string.Empty;
            string result;

            if (Equals(searchColumn, OKTMOColumn.Region))
            {
                if (string.IsNullOrEmpty(type))
                {
                    result =
                        results.FirstOrDefault(s => string.Equals(s, searchString, StringComparison.OrdinalIgnoreCase));
                }
                else
                {
                    result = results.FirstOrDefault(s => s.IndexOf(type, StringComparison.OrdinalIgnoreCase) >= 0);
                }
                if (string.IsNullOrEmpty(result))
                    result = results.First();
            }
            else
                result = results.First();


            return result ?? string.Empty;
        }

        public static string GetFullName(IEnumerable<OktmoRow> rows, string searchString, OKTMOColumn searchColumn)
        {
            var pattern = searchString.Trim() + "(\\b|$)";


            foreach (
                var s in
                    rows.Select(r => GetPropValueByName(r, searchColumn))
                        .Where(val => Regex.IsMatch(val, pattern, RegexOptions.IgnoreCase)))
            {
                return s;
            }

            return string.Empty;
        }

        public static string GetFullName(DataTable table, string searchString, OKTMOColumn searchColumn)
        {
            var pattern = searchString.Trim() + "(\\b|$)";

            if (table == null) return string.Empty;
            foreach (
                var row in
                    table.Rows.Cast<DataRow>()
                        .Where(
                            row =>
                                Regex.IsMatch(row[classificatorColumnDictionary[searchColumn] - 1].ToString(), pattern,
                                    RegexOptions.IgnoreCase)))
            {
                return row[classificatorColumnDictionary[searchColumn] - 1].ToString();
            }

            return string.Empty;
        }

        private static string GetPropValueByName(OktmoRow row, OKTMOColumn column)
        {
            var propName = GetOktmoRowPropName(column);

            return (row.GetType().GetProperty(propName).GetValue(row, null)??"").ToString();
        }

        private static string GetOktmoRowPropName(OKTMOColumn column)
        {
            var propName = "";

            var row = new OktmoRow();
            switch (column)
            {
                case OKTMOColumn.Subject:
                    propName = GetPropertyName(() => row.Subject);
                    break;
                case OKTMOColumn.Region:
                    propName = GetPropertyName(() => row.Region);
                    break;
                case OKTMOColumn.Settlement:
                    propName = GetPropertyName(() => row.Settlement);
                    break;
                case OKTMOColumn.NearCity:
                    propName = GetPropertyName(() => row.NearCity);
                    break;
                case OKTMOColumn.TypeOfNearCity:
                    propName = GetPropertyName(() => row.TypeOfNearCity);
                    break;
            }

            return propName;
        }

        private static string GetPropertyName<T>(Expression<Func<T>> expression)
        {
            MemberExpression body = (MemberExpression)expression.Body;
            return body.Member.Name;
        }

//        public void SetCustomTable(SearchParams searchParams)
//        {
//            var sourceTable = SubjectTable ?? table;
//
//            var currentCombination = new ColumnCombination {Subject = subjectName};
//
//            //Try to get cashed datatable
//            DataTable tmpTable;
//            if (Equals(searchParams.SearchColumn, OKTMOColumns.Region))
//            {
//                currentCombination.Region = searchParams.SearchString;
//
//                tmpTable = cashedCustomTables.FirstOrDefault(pair => pair.Key.Equals(currentCombination)).Value;
//                if (tmpTable != null)
//                {
//                    CustomDataTable = tmpTable;
//                    return;
//                }
//            }
//            else if (Equals(searchParams.SearchColumn, OKTMOColumns.NearCity))
//            {
//                currentCombination.NearCity = searchParams.SearchString;
//
//                tmpTable = cashedCustomTables.FirstOrDefault(pair => pair.Key.Equals(currentCombination)).Value;
//                if (tmpTable != null)
//                {
//                    CustomDataTable = tmpTable;
//                    return;
//                }
//            }
//
//            //Create new custom datatable
//            var searchColumn = GetExcelColumn(searchParams.SearchColumn) - 1;
//            var seachString = searchParams.SearchString;
//
//            CustomDataTable =
//                sourceTable.GetCustomDataTable(
//                    row =>
//                        string.Equals(ExcelLocationRow.ReplaceYo(row[searchColumn].ToString()), seachString,
//                            StringComparison.OrdinalIgnoreCase));
//
//            //Cash DataTable
//            if (cashedCustomTables.All(pair => !pair.Key.Equals(currentCombination)))
//                cashedCustomTables.Add(currentCombination, CustomDataTable);
//        }


        private class ColumnCombination : OktmoRow, IEquatable<ColumnCombination>
        {
            public bool Equals(ColumnCombination comparer)
            {
                if (!string.Equals(Subject, comparer.Subject))
                    return false;
                if (!string.Equals(Region, comparer.Region))
                    return false;
                if (!string.Equals(Settlement, comparer.Settlement))
                    return false;
                if (!string.Equals(NearCity, comparer.NearCity))
                    return false;
                if (!string.Equals(TypeOfNearCity, comparer.TypeOfNearCity))
                    return false;

                return true;
            }
        }
    }

    public struct ColumnNumbers
    {
        public ColumnNumbers(byte subject, byte region, byte settlement, byte nearcity, byte typeOfNearCity) : this()
        {
            Subject = subject;
            Region = region;
            Settlement = settlement;
            NearCity = nearcity;
            TypeOfNearCity = typeOfNearCity;
        }

        public byte Subject { get; set; }
        public byte Region { get; set; }
        public byte Settlement { get; set; }
        public byte NearCity { get; set; }
        public byte TypeOfNearCity { get; set; }
    }


    public class SearchParams
    {
        public SearchParams(string searchString, OKTMOColumn searchColumn)
        {
            SearchString = searchString;
            SearchColumn = searchColumn;
        }

        public string SearchString { get; private set; }
        public Enum SearchColumn { get; private set; }
    }
}