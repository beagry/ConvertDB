using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using Microsoft.Office.Interop.Excel;
using OfficeOpenXml;
using DataTable = System.Data.DataTable;

namespace Formater.SupportWorksheetsClasses
{

    public enum OKTMOColumns
    {
        //Название колонок с иерархией от общего к частному
        Subject = 1,
        Region = 2,
        Settlement = 3,
        NearCity = 4,
        TypeOfNearCity = 5
    }

    internal class OKTMOWorksheet
    {
//        private readonly Worksheet worksheet;
//        private readonly Worksheet regCentersWorksheet = null;
        private DataSet ds;
        private static long lastUsedRow;
        private string subjectName;

        private readonly DataTable table;
        private readonly DataTable regCTable = null;
        private const string regCWsName = "РегЦентры";

        private readonly DataTable mskTable;
        const string Moscow = "Москва";
        const string mskWsName = "территории Мск";
        
        private readonly DataTable spbTable;
        const string SPB = "Санкт-Петербург";
        const string spbWsName = "территории СПб";

        private class ColumnCombination:IEquatable<ColumnCombination>
        {
            public string Subject { get; set; } 
            public string Region { get; set; } 
            public string Settlement { get; set; } 
            public string NearCity { get; set; } 
            public string TypeOfNearCity { get; set; }

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

        private Dictionary<string, DataTable> reserveDataTables = new Dictionary<string, DataTable>();
        private Dictionary<string, DataTable> cashedSubjecTables = new Dictionary<string, DataTable>();
        private Dictionary<ColumnCombination, DataTable> cashedCustomTables = new Dictionary<ColumnCombination, DataTable>(); 

        public string RegCenter { get; set; }
        public DataTable SubjectTable { get; private set; }
        public DataTable CustomDataTable { get; private set; }


        private static readonly Dictionary<OKTMOColumns, byte> classificatorColumnDictionary = new Dictionary
            <OKTMOColumns, byte>
        {
            {OKTMOColumns.Subject, OKTMOColumnsFilter.Default.Subject},
            {OKTMOColumns.Region, OKTMOColumnsFilter.Default.Region},
            {OKTMOColumns.Settlement, OKTMOColumnsFilter.Default.Settlement},
            {OKTMOColumns.NearCity, OKTMOColumnsFilter.Default.NearCity},
            {OKTMOColumns.TypeOfNearCity, OKTMOColumnsFilter.Default.TypeOfNearCity},
        };


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

        public OKTMOWorksheet(DataSet ds, string mainWsName)
        {
            this.ds = ds;
            table = ds.Tables.Cast<DataTable>().First(t => t.TableName.Equals(mainWsName));
            lastUsedRow = table.Rows.Count;

            //Региональные центры
            regCTable = ds.Tables.Cast<DataTable>().FirstOrDefault(t => t.TableName.Equals(regCWsName));

            //Москва
            mskTable = ds.Tables.Cast<DataTable>().FirstOrDefault(t => t.TableName.Equals(mskWsName));

            //Санкт-Петербург
            spbTable = ds.Tables.Cast<DataTable>().FirstOrDefault(t => t.TableName.Equals(spbWsName));
        }

        public DataTable Table
        {
            get { return table; }
        }

        public bool StringMatchInColumn(DataTable table, string s, OKTMOColumns column)
        {
            if (String.IsNullOrEmpty(s)) return false;
            if (table == null)
                table = this.table;

            var res =
                table.Rows.Cast<DataRow>()
                    .Any(
                        row => row[GetExcelColumn(column) - 1] is string &&
                            string.Equals(DbToConvert.ReplaceYO(row[GetExcelColumn(column) - 1].ToString()), s,
                                StringComparison.OrdinalIgnoreCase));
            return res;
        }

        /// <summary>
        /// Аналог VlookUp in Excel
        /// </summary>
        /// <param name="contentColumn"></param>
        /// <param name="searchParams"></param>
        /// <returns></returns>
        public DataTable GetCustomDataTable(params SearchParams[] searchParams)
        {
            var dataTable = table.Copy(); 


            //Сортировка для поиска от общего к частному
            foreach (SearchParams @params in searchParams)
            {
                //Ищем все строки, в которых в ячейках по искомому столбцу строки содержат искомое значение
                var searchColumn = GetExcelColumn(@params.SearchColumn) - 1;
                var searchString = @params.SearchString;
                dataTable =
                    dataTable.GetCustomDataTable( //Метод создания новой таблицы по условию
                        row =>
                            String.Equals(row[searchColumn].ToString(), searchString,
                                StringComparison.CurrentCultureIgnoreCase)); //Полное совпадение
            }
            Debug.Assert(dataTable.Rows.Count > 0);
            //Из полученной таблицы достаём нужную нам колонку
            return dataTable;
        }


        public DataTable GetCustomDataTable(DataTable currenTable, params SearchParams[] searchParams)
        {
            if (String.IsNullOrEmpty(searchParams[0].SearchString)) return currenTable;
            //var result = new List<string>();
            DataTable dataTable = null;
            if (currenTable == null) currenTable = table;

            Debug.Assert(searchParams.Count() == 1); //Пока работаем только с одним критерием поиска

            //Сортировка для поиска от общего к частному
            //searchParams = searchParams.OrderBy(x => x.SearchColumn);

            foreach (SearchParams @params in searchParams)
            {

                //Ищем все строки, в которых в ячейках по искомому столбцу строки содержат искомое значение
                var searchColumn = GetExcelColumn(@params.SearchColumn) - 1;
                var searchString = @params.SearchString;

                //Try to use reserved tables
                if (@params.SearchColumn.Equals((Enum) OKTMOColumns.Subject) &&
                    reserveDataTables.ContainsKey(searchString)) //Если в поиске субъект
                    dataTable = reserveDataTables[searchString];

                    //Если в поиске муниципальное образование, то по одному образованию может быть несколько субъектом
                    //Смотрим что найденная таблица содержит тот же субъект что и ?
                else if (@params.SearchColumn.Equals((Enum) OKTMOColumns.Region) &&
                         reserveDataTables.ContainsKey(searchString))
                {
                    var column = searchColumn;
                    //Сравниваем зарезервированную таблицу с ижу имеющейся таблицей
                    //Важно чтобы субъект по искомому муниОбразованию совпадал
                    var reservedSubjects =
                        reserveDataTables[searchString].Rows.Cast<DataRow>()
                            .Select(row => row[GetExcelColumn(OKTMOColumns.Subject) - 1].ToString())
                            .Distinct()
                            .ToList();
                    var currentSubject =
                        currenTable.Rows.Cast<DataRow>()
                            .Where(row => row[column].ToString() == searchString)
                            .Select(row => row[GetExcelColumn(OKTMOColumns.Subject) - 1].ToString())
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
                                String.Equals(DbToConvert.ReplaceYO(row[searchColumn].ToString()), searchString,
                                    StringComparison.CurrentCultureIgnoreCase));
                    if ((@params.SearchColumn.Equals(OKTMOColumns.Subject) ||
                        @params.SearchColumn.Equals(OKTMOColumns.Region)) && !reserveDataTables.ContainsKey(searchString))
                        reserveDataTables.Add(searchString, dataTable);
                }
            }

            return dataTable;
        }

        public static byte GetExcelColumn(Enum searchColumn)
        {
            if (!classificatorColumnDictionary.ContainsKey((OKTMOColumns) searchColumn)) return 0;

            return (byte) (classificatorColumnDictionary[(OKTMOColumns) searchColumn]);
        }

        public List<string> GetContentByValue(OKTMOColumns contentColumn, string searchString, OKTMOColumns searchColumn,
            string searchString2, OKTMOColumns searchColumn2)
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
        /// Возвращает значение ячейки из выбранного столбца, с дополненным окончанием. Игнорирует регистр
        /// </summary>
        /// <param name="searchString">Искомое имя</param>
        /// <param name="searchColumn">Колонка поиска</param>
        /// <param name="type">Тип искомого наименования. Используется преимущественно для Муниципального Образования</param>
        /// <returns></returns>
        public string GetFullName(string searchString, OKTMOColumns searchColumn, string type = "" )
        {
            var pattern =searchString + "(\\b|$)";
            

            var results = table.Rows.Cast<DataRow>()
                .Where(row => Regex.IsMatch(DbToConvert.ReplaceYO(row[classificatorColumnDictionary[searchColumn] - 1].ToString()),pattern,RegexOptions.IgnoreCase))
                .Select(r => r[classificatorColumnDictionary[searchColumn] - 1].ToString()).Distinct().ToList();


            if (results.Count == 0) return string.Empty;
            string result;

            if (Equals(searchColumn, OKTMOColumns.Region))
            {
                if (string.IsNullOrEmpty(type))
                {
                    result = results.FirstOrDefault(s => string.Equals(s, searchString, StringComparison.OrdinalIgnoreCase));
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


            return result??String.Empty;
        }

        public static string GetFullName(DataTable table, string searchString, OKTMOColumns searchColumn)
        {
            var pattern = searchString + "(\\b|$)";

            if (table == null) return String.Empty;
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

            return String.Empty;
        }

        /// <summary>
        /// Возвращает региональный ценр переданного субъекта
        /// </summary>
        /// <param name="regionFullName">Название субъекта</param>
        /// <returns></returns>
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

            string res = row[regCenterTableColumn].ToString();
            res = res.Replace("город", "");
            res = res.Trim();

            return res;
        }

        /// <summary>
        /// Возвращает региональный ценрт, проверенный и приведенный к ОКТМО
        /// </summary>
        /// <param name="regionFullName"></param>
        /// <returns></returns>
        public string GetDefaultRegCenterFullName(string regionFullName, ref string cityName)
        {
            var regCenterName = GetDefaultRegCenter(regionFullName);
            if (string.IsNullOrEmpty(regCenterName)) return string.Empty;
            var rows =
                table.Rows.Cast<DataRow>()
                    .Where(
                        r =>
                            string.Equals(r[Columns.NearCity - 1].ToString(), regCenterName,
                                StringComparison.OrdinalIgnoreCase) &&
                            string.Equals(r[Columns.TypeOfNearCity - 1].ToString(), "город",
                                StringComparison.OrdinalIgnoreCase))
                    .ToList();
            var fullName = rows.Count == 1 ? rows[0][Columns.Region - 1].ToString() : String.Empty;
            //GetFullName("город " + regCenterName, OKTMOColumns.Region);
            cityName = regCenterName;
            return fullName;
        }

        public void SetCustomTable(SearchParams searchParams)
        {
            DataTable sourceTable = SubjectTable ?? table;

            ColumnCombination currentCombination = new ColumnCombination {Subject = subjectName};

            //Try to get cashed datatable
            DataTable tmpTable;
            if (Equals(searchParams.SearchColumn, OKTMOColumns.Region))
            {
                currentCombination.Region = searchParams.SearchString;

                tmpTable = cashedCustomTables.FirstOrDefault(pair => pair.Key.Equals(currentCombination)).Value;
                if (tmpTable != null)
                {
                    CustomDataTable = tmpTable;
                    return;
                }
            }
            else if (Equals(searchParams.SearchColumn, OKTMOColumns.NearCity))
            {
                currentCombination.NearCity = searchParams.SearchString;

                tmpTable = cashedCustomTables.FirstOrDefault(pair => pair.Key.Equals(currentCombination)).Value;
                if (tmpTable != null)
                {
                    CustomDataTable = tmpTable;
                    return;
                }
            }

            //Create new custom datatable
            var searchColumn = GetExcelColumn(searchParams.SearchColumn) - 1;
            var seachString = searchParams.SearchString;

            CustomDataTable =
                sourceTable.GetCustomDataTable(
                    row =>
                        string.Equals(DbToConvert.ReplaceYO(row[searchColumn].ToString()), seachString,
                            StringComparison.OrdinalIgnoreCase));

            //Cash DataTable
            if (cashedCustomTables.All(pair => !pair.Key.Equals(currentCombination)))
                cashedCustomTables.Add(currentCombination,CustomDataTable);
        }

    }

    internal struct ColumnNumbers
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


    internal class SearchParams
    {
        public string SearchString { get; private set; }
        public Enum SearchColumn { get; private set; }

        public SearchParams(string searchString, OKTMOColumns searchColumn)
        {
            SearchString = searchString;
            SearchColumn = searchColumn;
        }
    }
}
