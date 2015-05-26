using System.Collections.Generic;
using System.Linq;
using Converter.Template_workbooks;
using ExcelRLibrary;
using Microsoft.Office.Interop.Excel;
using OfficeOpenXml;
using DataTable = System.Data.DataTable;

namespace Converter
{
    /// <summary>
    /// Тупо создает новую книгу по переданным правилам
    /// Класс для объединения книг в шаблон на основе переданных правил
    /// </summary>
    /// <typeparam name="T">Книга-шаблон. Используется для создаения Excel обеъкта</typeparam>
    public class WorkbookTypifier<T> where T : TemplateWorkbook, new ()
    {
        public ICollection<string> WorkbooksPaths { get; set; }

        /// <summary>
        /// Правила для объединения книг
        /// </summary>
        public Dictionary<string, List<string>> RulesDictionary { get; set; }



        public WorkbookTypifier(Dictionary<string, List<string>> rulesDictionary, ICollection<string> workbooksPaths)
        {
            RulesDictionary = rulesDictionary;
            this.WorkbooksPaths = workbooksPaths;
        }

        public WorkbookTypifier()
        {
            RulesDictionary = new Dictionary<string, List<string>>();
            WorkbooksPaths = new List<string>();
            
        }




        /// <summary>
        /// Метод возвращает единую книгу, солженную из переданныхх книг по переданным правилам
        /// </summary>
        /// <param name="workbooksPaths"></param>
        /// <returns></returns>
        public ExcelPackage CombineToSingleWorkbook()
        {
            var result = new ExcelPackage();
            var resultWS =  result.Workbook.Worksheets.Add("Combined");


            //подготовить конечный лист
            var templateHead = new T().TemplateColumns.ToDictionary(k => k.Index, v => v.CodeName);
            resultWS.WriteHead(templateHead);

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