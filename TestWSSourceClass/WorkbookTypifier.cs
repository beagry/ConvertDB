using System.Collections.Generic;
using System.Linq;
using Converter.Template_workbooks;
using ExcelRLibrary;
using Microsoft.Office.Interop.Excel;

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
        public Workbook CombineToSingleWorkbook()
        {
            var helper = new ExcelHelper();

            //создать пустую книгу
            var newWb = helper.CreateNewWorkbook();
            var ws = newWb.Worksheets[1] as Worksheet;
            
            //оформить шапку
            var templateHead = new T().TemplateColumns.ToDictionary(k => k.Index, v => v.CodeName);
            ws.WriteHead(templateHead);

            var wsWriter = new WorksheetFiller(ws, RulesDictionary);
            
            //поочередно записываем книги из списка книги из списка
            foreach (var openWs in helper.GetWorkbooks(WorkbooksPaths).Select(wb => wb.Worksheets[1]).Cast<Worksheet>())
                wsWriter.InsertWorksheet(openWs);

            return newWb;
        }        
    }
}