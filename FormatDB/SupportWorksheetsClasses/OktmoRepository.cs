using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Data;
using System.Diagnostics;
using System.Linq;
using System.Linq.Expressions;
using System.Text.RegularExpressions;
using AutoMapper;
using AutoMapper.QueryableExtensions;
using ExcelRLibrary;
using ExcelRLibrary.SupportEntities.Oktmo;
using NLog;
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
        private static readonly Dictionary<OKTMOColumn, byte> ClassificatorColumnDictionary = new Dictionary
            <OKTMOColumn, byte>
        {
            {OKTMOColumn.Subject, OKTMOColumnsFilter.Default.Subject},
            {OKTMOColumn.Region, OKTMOColumnsFilter.Default.Region},
            {OKTMOColumn.Settlement, OKTMOColumnsFilter.Default.Settlement},
            {OKTMOColumn.NearCity, OKTMOColumnsFilter.Default.NearCity},
            {OKTMOColumn.TypeOfNearCity, OKTMOColumnsFilter.Default.TypeOfNearCity}
        };

        private readonly Logger logger = LogManager.GetCurrentClassLogger();
        private ICollection<OktmoRow> oktmoRows;
        private readonly DataTable regCTable;

        public OKTMORepository(DataSet ds, string mainWsName)
        {
            var table = ds.Tables.Cast<DataTable>().First(t => t.TableName.Equals(mainWsName));
            regCTable = ds.Tables.Cast<DataTable>().FirstOrDefault(t => t.TableName.EqualNoCase("РегЦентры"));

            if (regCTable == null)
            {
                logger.Warn(
                    "Не найдена база с Региональными центрами, заполнение региональных центров не будет выполнено.");
            }

            oktmoRows = table.Rows.Cast<DataRow>().Select(r => new OktmoRow
            {
                Subject = (r[GetExcelColumn(OKTMOColumn.Subject) - 1] ?? "").ToString(),
                Region = (r[GetExcelColumn(OKTMOColumn.Region) - 1] ?? "").ToString(),
                Settlement = (r[GetExcelColumn(OKTMOColumn.Settlement) - 1] ?? "").ToString(),
                NearCity = (r[GetExcelColumn(OKTMOColumn.NearCity) - 1] ?? "").ToString(),
                TypeOfNearCity = (r[GetExcelColumn(OKTMOColumn.TypeOfNearCity) - 1] ?? "").ToString()
            }).ToList().AsReadOnly();
        }

        public OKTMORepository()
        {
            oktmoRows = new List<OktmoRow>();
        }

        public void InitializeFromDb()
        {
            var db = new OktmoContext();
            try
            {
                Mapper.AddProfile<OktmoProfile>();
                oktmoRows = db.OktmoRows.Project().To<OktmoRow>().ToList();
            }
            catch (Exception)
            {

                throw;
            }
            finally
            {
                db.Dispose();
            }
            
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

        public string GetDefaultRegCenter(string regionFullName)
        {
            if (string.IsNullOrEmpty(regionFullName)) return string.Empty;
            if (regCTable == null) return string.Empty;
            const int regionTableColumn = 0;
            const  int regCenterTableColumn = 1;

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
            var regCenterName = GetDefaultRegCenter(regionFullName);
            if (string.IsNullOrEmpty(regCenterName)) return string.Empty;

            var regCentSpec =
                new ExpressionSpecification<OktmoRow>(
                    row => string.Equals(row.NearCity, regCenterName, StringComparison.OrdinalIgnoreCase));
            var cityRegCentSpec =
                new ExpressionSpecification<OktmoRow>(
                    row => string.Equals(row.TypeOfNearCity, "город", StringComparison.OrdinalIgnoreCase));
            var spec = regCentSpec.And(cityRegCentSpec);

            var rows = oktmoRows.Where(r => spec.IsSatisfiedBy(r)).DistinctBy(r => r.Region).ToList();
            if (rows.Count() != 1) return string.Empty;
            var singleRow = rows.Single();

            var fullName = singleRow != null ? singleRow.Region : string.Empty;

            cityName = regCenterName;
            return fullName;
        }

        public IEnumerable<OktmoRow> GetSubjectRows(string subj)
        {
            var spec = new SubjectSpecification(subj);
            return oktmoRows.Where(oktmoRow => spec.IsSatisfiedBy(oktmoRow));
        }

        public IEnumerable<OktmoRow> GetRowsByRegion(string reg)
        {
            var spec =
                new ExpressionSpecification<OktmoRow>(
                    row => string.Equals(row.Region, reg, StringComparison.OrdinalIgnoreCase));
            return oktmoRows.Where(oktmoRow => spec.IsSatisfiedBy(oktmoRow));
        }

        public IEnumerable<OktmoRow> GetRowsBySettlement(string sett)
        {
            var spec =
                new ExpressionSpecification<OktmoRow>(
                    row => string.Equals(row.Settlement, sett, StringComparison.OrdinalIgnoreCase));
            return oktmoRows.WhoSatisfySpec(spec);
        }

        public IEnumerable<OktmoRow> GetRowsByNearCity(string city)
        {
            var spec =
                new ExpressionSpecification<OktmoRow>(
                    row => string.Equals(row.NearCity, city, StringComparison.OrdinalIgnoreCase));
            return oktmoRows.Where(oktmoRow => spec.IsSatisfiedBy(oktmoRow));
        }

        public IEnumerable<OktmoRow> GetRowsBySearchParams(OktmoRow searchParamsRow)
        {
            var spec = new OktmoSpecifications(searchParamsRow);
            return oktmoRows.Where(oktmoRow => spec.IsSatisfiedBy(oktmoRow));
        }


        private static byte GetExcelColumn(Enum searchColumn)
        {
            if (!ClassificatorColumnDictionary.ContainsKey((OKTMOColumn) searchColumn)) return 0;

            return ClassificatorColumnDictionary[(OKTMOColumn) searchColumn];
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
            var pattern = "(\\b|^)" +  searchString.Trim() + "(\\b|$)";


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
            var pattern = "(\\b|^)" + searchString.Trim() + "(\\b|$)";


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
                                Regex.IsMatch(row[ClassificatorColumnDictionary[searchColumn] - 1].ToString(), pattern,
                                    RegexOptions.IgnoreCase)))
            {
                return row[ClassificatorColumnDictionary[searchColumn] - 1].ToString();
            }

            return string.Empty;
        }

        private static string GetPropValueByName(OktmoRow row, OKTMOColumn column)
        {
            var propName = GetOktmoRowPropName(column);

            return (row.GetType().GetProperty(propName).GetValue(row, null) ?? "").ToString();
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
            var body = (MemberExpression) expression.Body;
            return body.Member.Name;
        }

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