using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Converter.Models;
using Converter.Template_workbooks;
using Converter.Tools;
using ExcelRLibrary;
using Microsoft.Office.Interop.Excel;
using DataTable = System.Data.DataTable;

namespace Converter
{

    /// <summary>
    /// Помощник в анализе книг
    /// В результате выдаёт список столбцов наиболее подходящих к шаблону
    /// </summary>
    public class WorkbooksAnalyzier: IDisposable
    {
        private readonly XlTemplateWorkbookType wbType;
        private readonly ExcelHelper excelHelper;
        private TemplateWorkbook templateWorkbook;


        //Result properties
        public Dictionary<string,List<string>> ComparedColumns { get; private set; }
        public List<WorksheetInfo> WorksheetsInfos { get; set; }


        public WorkbooksAnalyzier(XlTemplateWorkbookType workbookType)
        {
            WorksheetsInfos = new List<WorksheetInfo>();
            excelHelper = new ExcelHelper();
            wbType = workbookType;
            templateWorkbook = wbType.GetWorkbook();
            CreateResultDict();
        }


        public async Task CheckWorkbooksAsync(IEnumerable<string> wbPaths)
        {
            foreach (var wbPath in wbPaths)
            {
                var path = wbPath;
                await Task.Run(()=>CheckWorkbook(path));
            }
        }

        public void CheckWorkbook(string path, string wsName = "1")
        {
            var fi = new FileInfo(path);
            var reader = new ExcelReader();
            var ds = reader.ReadExcelFile(fi.FullName);
            var dt = ds.Tables.Cast<DataTable>().First();
            if (dt == null) return;


            //Создаем модель рабочего листа
            WorksheetsInfos.Add(new WorksheetInfo(dt){Workbook = new SelectedWorkbook(fi.FullName)});

            //Анализируем содержание рабочего листа
            var sourceWs = new SourceWs(dt, templateWorkbook);
            sourceWs.CheckColumns();
            var result = sourceWs.ResultDictionary;

            //Add to compareResultDictionary
            foreach (var keyPair in result)
            {
                var templateColumnName = keyPair.Key;
                var comparedColumnNames = keyPair.Value;

                if (!ComparedColumns.ContainsKey(templateColumnName))
                    continue;
//                    ComparedColumns.Add(templateColumnName, new List<string>());

                comparedColumnNames.ForEach(s =>
                {
                    if (!ComparedColumns[templateColumnName].Contains(s))
                        ComparedColumns[templateColumnName].Add(s);
                });
            }
        }

        public void CheckWorkbook(Workbook wb, byte wsIndex = 1)
        {
            Worksheet ws;
            try
            {
                ws = (Worksheet) wb.Worksheets[wsIndex];
                try
                {
                    ws.ShowAllData();

                }
                catch (Exception)
                {
                    //ignored
                }
            }
            catch (Exception e)
            {
                throw e;
            }

            //Создаем модель рабочего листа
            WorksheetsInfos.Add(new WorksheetInfo(ws){Workbook = new SelectedWorkbook(wb.FullName) });

            //Анализируем содержание рабочего листа
            var sourceWs = new SourceWs(ws,templateWorkbook);
            sourceWs.CheckColumns();
            var result = sourceWs.ResultDictionary;

            //Add to compareResultDictionary
            foreach (var keyPair in result)
            {
                var templateColumnName = keyPair.Key;
                var comparedColumnNames = keyPair.Value;

                if (!ComparedColumns.ContainsKey(templateColumnName))
                    ComparedColumns.Add(templateColumnName, new List<string>());

                comparedColumnNames.ForEach(s =>
                {
                    if (!ComparedColumns[templateColumnName].Contains(s))
                        ComparedColumns[templateColumnName].Add(s);
                });
                
            }
        }

        private void CreateResultDict()
        {
            ComparedColumns = templateWorkbook.TemplateColumns.ToDictionary(j => j.CodeName, j2 => new List<string>());
        }


        public void Dispose()
        {
            if (excelHelper != null)
                excelHelper.Dispose();
        }
    }
}
