using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Runtime.InteropServices;
using Formater.SupportWorksheetsClasses;
using Microsoft.Office.Interop.Excel;
using DataTable = System.Data.DataTable;

namespace Formater
{
    static class ExcelExtensions
    {
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

        public static readonly Color BadColor = Color.Crimson;
        public static readonly Color GoodColor = Color.Aquamarine;
        public static readonly Color Clear = Color.Transparent;
    }
}