using System;
using System.Collections.Generic;
using System.Linq;
using ExcelRLibrary.TemplateWorkbooks;
using Microsoft.Office.Interop.Excel;

namespace Converter.Template_workbooks
{
    /// <summary>
    ///     Абстрактный класс шаблонной книги
    /// </summary>
    public class TemplateWorkbook
    {
        protected List<JustColumn> Columns;

        public string UnUsedColumnCode
        {
            get { return "UNUS"; }
        }

        public IEnumerable<JustColumn> TemplateColumns
        {
            get { return Columns; }
        }

        public int GetColumnByCode(string name)
        {
            var column = 0;
            var firstOrDefault = Columns.FirstOrDefault(x => x.CodeName == name);
            if (firstOrDefault != null)
                column = firstOrDefault.Index;
            return column;
        }

        public static Workbook GetTemplateWorkbook()
        {
            throw new Exception("Метод не готов!");
            //Excel.Workbook workbook = (new Excel.Application()).Workbooks.Add();
            //return workbook;
        }
    }

    public class WSType
    {
        public List<string> Heads { get; set; }
        public int GroupNumber { get; set; }
    }
}