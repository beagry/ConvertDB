using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using Converter.Models;
using Converter.Template_workbooks;
using Converter.Template_workbooks.EFModels;
using ExcelRLibrary;
using ExcelRLibrary.TemplateWorkbooks;

namespace Converter
{
    /// <summary>
    ///     Помощник в анализе книг
    ///     В результате выдаёт список столбцов наиболее подходящих к шаблону
    /// </summary>
    public class WorkbooksAnalyzier
    {
        private readonly XlTemplateWorkbookType wbType;
        private Dictionary<JustColumn, List<string>> comparedColumns;

        public WorkbooksAnalyzier(XlTemplateWorkbookType workbookType)
        {
            WorksheetsInfos = new List<WorksheetInfo>();
            wbType = workbookType;
        }

        /// <summary>
        ///     Result of CheckWorkbook(s) Method
        /// </summary>
        public Dictionary<JustColumn, List<string>> ComparedColumns
        {
            get { return comparedColumns ?? (comparedColumns = CreateResultDict()); }
        }

        /// <summary>
        ///     Info about worksheets of WB
        /// </summary>
        public List<WorksheetInfo> WorksheetsInfos { get; private set; }

        /// <summary>
        ///     Метод пытается найти соотвествующие колонки для шаблонной книги
        ///     Резултат процедуры будет находиться в переменной ComparedColumns
        /// </summary>
        /// <param name="wbPaths"></param>
        /// <returns></returns>
        public void CheckWorkbooks(IEnumerable<string> wbPaths)
        {
            try
            {
                foreach (var wbPath in wbPaths)
                {
                    var path = wbPath;
                    CheckWorkbook(path);
                }
            }
            catch (IOException)
            {
                throw new IOException("Не удалось прочитать файлы.");
            }
        }

        private void CheckWorkbook(string path)
        {
            var fi = new FileInfo(path);
            var reader = new ExcelReader();
            DataSet ds = reader.ReadExcelFile(fi.FullName);
            if (ds == null) throw new IOException("Не удалось прочитать файл.");

            var dt = ds.Tables.Cast<DataTable>().First();
            if (dt == null) return;


            //Создаем модель рабочего листа
            WorksheetsInfos.Add(new WorksheetInfo(dt) {Workbook = new SelectedWorkbook(fi.FullName)});

            //Анализируем содержание рабочего листа
            var sourceWs = new SourceWs(dt, wbType);
            sourceWs.CheckColumns();
            var result = sourceWs.ResultDictionary;

            //Add to compareResultDictionary
            foreach (var keyPair in result)
            {
                var templateColumnName = keyPair.Key;
                var comparedColumnNames = keyPair.Value;

                if (!ComparedColumns.Keys.Any(j => j.CodeName.Equals(templateColumnName)))
                    continue;

                var list = ComparedColumns.First(pair => pair.Key.CodeName.Equals(templateColumnName)).Value;

                comparedColumnNames.ForEach(s =>
                {
                    if (!list.Contains(s))
                        list.Add(s);
                });
            }
        }

        private Dictionary<JustColumn, List<string>> CreateResultDict()
        {
            var wb = UnitOfWorkSingleton.UnitOfWork.TemplateWbsRespository.GetTypedWorkbook(wbType);
            var columns = wb.Columns.Select(c => new JustColumn(c.CodeName, c.Name, c.ColumnIndex)).ToList();
            return columns.ToDictionary(j => j, j2 => new List<string>());
        }
    }
}