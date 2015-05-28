﻿using System;
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
using Converter.Template_workbooks.EFModels;
using Converter.Tools;
using ExcelRLibrary;
using ExcelRLibrary.TemplateWorkbooks;
using Microsoft.Office.Interop.Excel;
using DataTable = System.Data.DataTable;
using TemplateWorkbook = Converter.Template_workbooks.TemplateWorkbook;

namespace Converter
{

    /// <summary>
    /// Помощник в анализе книг
    /// В результате выдаёт список столбцов наиболее подходящих к шаблону
    /// </summary>
    public class WorkbooksAnalyzier
    {
        private readonly XlTemplateWorkbookType wbType;

        /// <summary>
        ///     Result of CheckWorkbook(s) Method
        /// </summary>
        public Dictionary<JustColumn,List<string>> ComparedColumns { get; private set; }

        /// <summary>
        ///     Info about worksheets of WB
        /// </summary>
        public List<WorksheetInfo> WorksheetsInfos { get; private set; }


        public WorkbooksAnalyzier(XlTemplateWorkbookType workbookType)
        {
            WorksheetsInfos = new List<WorksheetInfo>();
            wbType = workbookType;
            CreateResultDict();
        }


        /// <summary>
        ///     Метод пытается найти соотвествующие колонки для шаблонной книги
        ///     Резултат процедуры будет находиться в переменной ComparedColumns
        /// </summary>
        /// <param name="wbPaths"></param>
        /// <returns></returns>
        public void CheckWorkbooks(IEnumerable<string> wbPaths)
        {
            foreach (var wbPath in wbPaths)
            {
                var path = wbPath;
                CheckWorkbook(path);
            }
        }



        private void CheckWorkbook(string path)
        {
            var fi = new FileInfo(path);
            var reader = new ExcelReader();
            var ds = reader.ReadExcelFile(fi.FullName);
            var dt = ds.Tables.Cast<DataTable>().First();
            if (dt == null) return;


            //Создаем модель рабочего листа
            WorksheetsInfos.Add(new WorksheetInfo(dt){Workbook = new SelectedWorkbook(fi.FullName)});

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

        private void CreateResultDict()
        {
            var wb = TemplateWbsRepository.Context.TemplateWorkbooks.First(w => w.WorkbookType == wbType);
            var columns = wb.Columns.Select(c =>  new JustColumn(c.CodeName,c.Name,c.ColumnIndex));
            ComparedColumns = columns.ToDictionary(j => j, j2 => new List<string>());
        }

    }
}