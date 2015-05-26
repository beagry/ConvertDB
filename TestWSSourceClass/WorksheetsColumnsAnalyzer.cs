using System;
using System.Collections.Generic;
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

namespace Converter
{

    /// <summary>
    /// Помощник в анализе книг
    /// В результате выдаёт список столбцов наиболее подходящих к шаблону
    /// </summary>
    public class WorkbooksAnalyzier: IDisposable
    {
        private readonly XlTemplateWorkbookTypes wbType;
        private readonly ExcelHelper excelHelper;
        private TemplateWorkbook templateWorkbook;


        //Result properties
        public Dictionary<string,List<string>> ComparedColumns { get; set; }
        public List<WorksheetInfo> WorksheetsInfos { get; set; }


        public WorkbooksAnalyzier(XlTemplateWorkbookTypes workbookType)
        {
            WorksheetsInfos = new List<WorksheetInfo>();
            excelHelper = new ExcelHelper();
            wbType = workbookType;
            templateWorkbook = wbType.GetWorkbook();
            CreateResultDict();
        }


        public async Task CheckWorkbooksAsync(IEnumerable<string> wbPaths)
        {
            var app = ExcelHelper.App;
            app.DisplayAlerts = false;

            await Task.Run(() =>
            {
                try
                {
                    wbPaths.ForEach(s =>
                    {
                        var wb = excelHelper.GetWorkbook(s);
                        app.Visible = false;
                        CheckWorkbook(wb);
                        wb.Close();
                        Marshal.FinalReleaseComObject(wb);
                    });
                }
                catch (Exception)
                {
                    app.DisplayAlerts = true;
                    app.Visible = true;
                    throw;
                }
            });
            
            
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
                    throw;
                }
            }
            catch (Exception e)
            {
                throw;
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
