using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Excel;
using DataTable = System.Data.DataTable;

namespace Formater.SupportWorksheetsClasses
{
    public class VGTRepository
    {
        private readonly DataTable table;
        private const byte CityNameExcelColumn = 1;
        private const byte RegionNameExcelColumn = 4;
        private const byte TerritoryExcelColumn = 3;

        public VGTRepository(DataTable table)
        {
            this.table = table;
        }

        /// <summary>
        /// Проверяет есть ли переданный район в справочнике ВГТ. Игнорирует регистр
        /// </summary>
        /// <param name="s"></param>
        /// <returns>True если переданный район есть в справочнике</returns>
        public bool TerritotyExists(string s)
        {
            return
                table.Rows.Cast<DataRow>()
                    .Select(row => row[TerritoryExcelColumn - 1])
                    .OfType<string>()
                    .Any(s2 => String.Equals(s.Trim(), s2.Trim(), StringComparison.OrdinalIgnoreCase));
        }

        /// <summary>
        /// Возвращает True если в справочнике присутствует комбинация МунОбразования и Района
        /// </summary>
        /// <param name="region"></param>
        /// <param name="territory"></param>
        /// <returns></returns>
        public bool CombinationExists(string region, string territory)
        {
            var result =
                table.Rows.Cast<DataRow>()
                    .Where(r => r[RegionNameExcelColumn - 1] is string && r[TerritoryExcelColumn - 1] is string)
                    .Any(
                        r =>
                            String.Equals(r[RegionNameExcelColumn - 1].ToString(), region, StringComparison.OrdinalIgnoreCase) &&
                            String.Equals(r[TerritoryExcelColumn - 1].ToString(), territory,
                                StringComparison.OrdinalIgnoreCase));

            return result;
        }

        /// <summary>
        /// Возвращает наименование города из справочника, если район уникальный
        /// </summary>
        /// <param name="s">Наименование района</param>
        /// <returns></returns>
        public string GetCityByTerritory(string s)
        {
            var rows = GetCitiesListByTerritory(s);
            return rows.Count == 1 ? rows[0] : string.Empty;
        }

        /// <summary>
        /// Возвращает список городов, в которых присутствует район с переданным названием
        /// </summary>
        /// <param name="s">полное наименование района</param>
        /// <returns></returns>
        public List<string> GetCitiesListByTerritory(string s)
        {
            var cities =
                table.Rows.Cast<DataRow>()
                    .Where(
                        row =>
                            String.Equals(row[TerritoryExcelColumn - 1].ToString(), s,
                                StringComparison.OrdinalIgnoreCase))
                    .Select(row => row[CityNameExcelColumn - 1].ToString())
                    .Distinct()
                    .ToList();
            return cities;
        }

        private bool disposed = false;
        public void Dispose()
        {
            if (disposed) return;
            table.Dispose();
            disposed = true;
        }
    }
}
