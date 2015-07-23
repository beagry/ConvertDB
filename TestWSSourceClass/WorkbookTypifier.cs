using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using Converter.Template_workbooks;
using Converter.Template_workbooks.EFModels;
using ExcelRLibrary;
using OfficeOpenXml;

namespace Converter
{
    /// <summary>
    ///     Класс для объединения книг в шаблон на основе переданных правил
    /// </summary>
    public class WorkbookTypifier
    {
        public WorkbookTypifier()
        {
            RulesDictionary = new Dictionary<string, List<string>>();
            WorkbooksPaths = new List<string>();
            WorkbookType = XlTemplateWorkbookType.LandProperty;
        }

        public TemplateWorkbook TemplateWorkbook { get; set; }
        public string BaseWbPath { private get; set; }
        public ICollection<string> WorkbooksPaths { get; set; }

        /// <summary>
        ///     Правила для объединения книг
        /// </summary>
        public Dictionary<string, List<string>> RulesDictionary { get; set; }

        public XlTemplateWorkbookType WorkbookType { get; set; }

        /// <summary>
        ///     Метод возвращает единую книгу, солженную из переданныхх книг по переданным правилам
        /// </summary>
        /// <returns></returns>
        public ExcelPackage CombineToSingleWorkbook()
        {
            ExcelPackage result = new ExcelPackage();
            ExcelWorksheet resultWS = null;
            if (!string.IsNullOrEmpty(BaseWbPath))
            {
                result = new ExcelPackage(new FileInfo(BaseWbPath));
                resultWS = result.Workbook.Worksheets.First();
            }
            
            if (resultWS == null)
            {
                result = new ExcelPackage();
                resultWS = result.Workbook.Worksheets.Add("Combined");


                var columns = TemplateWorkbook.Columns.Select(c => new { Index = c.ColumnIndex, c.Name, Code = c.CodeName }).ToList();
                resultWS.WriteHead(columns.ToDictionary(k => k.Index, v => v.Code), 1);
                resultWS.WriteHead(columns.ToDictionary(k => k.Index, v => v.Name), 2);
            }

            var wsWriter = new WorksheetFiller(resultWS, RulesDictionary);


            var reader = new ExcelReader();
            foreach (
                var dt in
                    WorkbooksPaths.Select(p => reader.ReadExcelFile(p))
                        .Select(ds => ds.Tables.Cast<DataTable>().First()))
            {
                wsWriter.AppendDataTable(dt);
                //бug ошибка при сохранении
                //Можно откатить на Ctrl + Z
//                result.Save();
            }

            return result;
        }
    }
}