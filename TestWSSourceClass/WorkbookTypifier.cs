using System.Collections.Generic;
using System.Data;
using System.Linq;
using Converter.Template_workbooks;
using Converter.Template_workbooks.EFModels;
using ExcelRLibrary;
using OfficeOpenXml;

namespace Converter
{
    /// <summary>
    ///     Тупо создает новую книгу по переданным правилам
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

        public ICollection<string> WorkbooksPaths { get; set; }

        /// <summary>
        ///     Правила для объединения книг
        /// </summary>
        public Dictionary<string, List<string>> RulesDictionary { get; set; }

        public XlTemplateWorkbookType WorkbookType { get; set; }

        /// <summary>
        ///     Метод возвращает единую книгу, солженную из переданныхх книг по переданным правилам
        /// </summary>
        /// <param name="workbooksPaths"></param>
        /// <returns></returns>
        public ExcelPackage CombineToSingleWorkbook()
        {
            var result = new ExcelPackage();
            var resultWS = result.Workbook.Worksheets.Add("Combined");

            //подготовить конечный лист
            var wbRepo = new TemplateWbsRespository();

            var wb = wbRepo.GetTypedWorkbook(WorkbookType);
            var columns = wb.Columns.Select(c => new {Index = c.ColumnIndex, c.Name, Code = c.CodeName}).ToList();
            resultWS.WriteHead(columns.ToDictionary(k => k.Index, v => v.Code), 1);
            resultWS.WriteHead(columns.ToDictionary(k => k.Index, v => v.Name), 2);

            var wsWriter = new WorksheetFiller(resultWS, RulesDictionary);


            var reader = new ExcelReader();
            foreach (
                var dt in
                    WorkbooksPaths.Select(p => reader.ReadExcelFile(p))
                        .Select(ds => ds.Tables.Cast<DataTable>().First()))
            {
                wsWriter.AppendDataTable(dt);
            }

            return result;
        }
    }
}