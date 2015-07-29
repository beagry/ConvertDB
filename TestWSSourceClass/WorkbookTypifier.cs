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
            var result = new ExcelPackage();
            ExcelWorksheet resultWS = null;
            WorksheetFiller wsWriter = null;
            var reader = new ExcelReader();

            if (!string.IsNullOrEmpty(BaseWbPath) && File.Exists(BaseWbPath))
            {
                if (Path.GetExtension(BaseWbPath) == ".csv")
                {
                    result = new ExcelPackage();
                    var data = reader.ReadExcelFile(BaseWbPath).Tables.Cast<DataTable>().First();
                    resultWS = result.Workbook.Worksheets.Add(data.TableName);
                    resultWS.Cells["A1"].LoadFromDataTable(data , true);
                    data.Dispose();
                }
                else
                {
                    result = new ExcelPackage(new FileInfo(BaseWbPath));
                    resultWS = result.Workbook.Worksheets.First();
                }

                wsWriter = new WorksheetFiller(resultWS, RulesDictionary);
            }
            
            if (resultWS == null)
            {
                result = new ExcelPackage();
                resultWS = result.Workbook.Worksheets.Add("Combined");


                var columns = TemplateWorkbook.Columns.Select(c => new { Index = c.ColumnIndex, c.Name, Code = c.CodeName }).ToList();
                resultWS.WriteHead(columns.ToDictionary(k => k.Index, v => v.Code), 2);
                resultWS.WriteHead(columns.ToDictionary(k => k.Index, v => v.Name), 1);
                wsWriter = new WorksheetFiller(resultWS, RulesDictionary)
                {
                    HeadsDictionary = columns.ToDictionary(k => k.Index, v => v.Code)
                };
            }


            
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