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
using TemplateWorkbook = Converter.Template_workbooks.EFModels.TemplateWorkbook;

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
        public static TemplateWorkbook TemplateWb { get; set; }
        private readonly string mainWbPath;

        public WorkbooksAnalyzier(XlTemplateWorkbookType workbookType):this()
        {
            wbType = workbookType;
        }

        public WorkbooksAnalyzier(string mainBasePath):this()
        {
            mainWbPath = mainBasePath;
            wbType = XlTemplateWorkbookType.Custom;
        }

        public WorkbooksAnalyzier()
        {
            WorksheetsInfos = new List<WorksheetInfo>();
        }


        /// <summary>
        ///     Результат работы сравнения переданных книг и книги шаблона.
        ///     Даныне в формате Key = Столбец из книги шаблона, Value = Подошедшие столбцы из переданных книг
        /// </summary>
        public Dictionary<JustColumn, List<string>> ComparedColumns
        {
            get { return comparedColumns ?? (comparedColumns = CreateResultDict()); }
        }

        /// <summary>
        ///     Краткая инфомрация о проверенных книгах
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
            //TODO при создании книги-шаблона на основе базу подтягивать критерии поиска для столбцов, совпадающих с столбцами в базе
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

        public void CheckWorkbook(string path)
        {
            var fi = new FileInfo(path);
            var reader = new ExcelReader();
            DataSet ds = reader.ReadExcelFile(fi.FullName);

            var dt = ds.Tables.Cast<DataTable>().First();
            if (dt == null) return;


            //Создаем модель рабочего листа
            WorksheetsInfos.Add(new WorksheetInfo(dt) {Workbook = new SelectedWorkbook(fi.FullName)});

            if (TemplateWb == null)
            {
                InitialTempateWb();
            }

            //Анализируем содержание рабочего листа
            var sourceWs = new SourceWs(dt, TemplateWb);
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

        private void InitialTempateWb()
        {
            if (wbType == XlTemplateWorkbookType.Custom)
            {
                CreateTemppateWbFromExcel();
                return;
            }

            TemplateWb =
                UnitOfWorkSingleton.UnitOfWork.TemplateWbsRespository.GetObjectsList()
                    .First(w => w.WorkbookType == wbType);

        }

        private void CreateTemppateWbFromExcel()
        {
            var reader = new ExcelReader();
            var ds = reader.GetWsStructs(mainWbPath);
            var dt = ds.Tables.Cast<DataTable>().First();

            var columns = dt.Columns.Cast<DataColumn>().Select(dc => new TemplateColumn
            {
                ColumnIndex = dt.Columns.IndexOf(dc) + 1,
                Name = dc.ColumnName,
                CodeName = dc.ColumnName
            }).ToList();

            TemplateWb = new TemplateWorkbook {WorkbookType = wbType,Columns = columns};

            //Теперь попытаемся добавить колонки для поиска
            var templatedColumns =
                UnitOfWorkSingleton.UnitOfWork.TemplateWbsRespository.Context.TemplateColumns.ToList();
            foreach (var column in columns)
            {
                var column1 = column;
                var done = false;
                foreach (
                    var comparedColumn in
                        templatedColumns.Where(
                            c => c.CodeName.EqualNoCase(column1.CodeName) || c.Name.EqualNoCase(column1.Name)))
                {
                    column.BindedColumns = comparedColumn.BindedColumns;
                    column.SearchCritetias = comparedColumn.SearchCritetias;
                    done = true;
                }
                if (done) continue;

                foreach (
                    var comparedColumn in
                        templatedColumns.Where(
                            c =>
                                c.BindedColumns.Any(
                                    bc => bc.Name.EqualNoCase(column1.Name) || bc.Name.EqualNoCase(column1.CodeName))))
                {
                    column1.BindedColumns = comparedColumn.BindedColumns;
                    column1.BindedColumns.Add(new BindedColumn{Name = comparedColumn.Name});
                    column1.BindedColumns.Add(new BindedColumn{Name = comparedColumn.CodeName});
                }
            }
        }

        private Dictionary<JustColumn, List<string>> CreateResultDict()
        {
            var columns = TemplateWb.Columns.Select(c => new JustColumn(c.CodeName, c.Name, c.ColumnIndex)).ToList();
            return columns.ToDictionary(j => j, j2 => new List<string>());
        }
    }
}