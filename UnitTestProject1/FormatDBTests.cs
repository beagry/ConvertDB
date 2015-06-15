using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Converter.Template_workbooks;
using ExcelRLibrary;
using Formater;
using Formater.SupportWorksheetsClasses;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using UI;

namespace UnitTestProject1
{
    [TestClass]
    public class FormatDbTests
    {
        [TestMethod]
        public void RunWindow()
        {
            var w = new FormatWbWindow();
            w.ShowDialog();
            Assert.AreEqual(true, true);
        }

        [TestMethod]
        public void CheckRowsWithParaps()
        {
            const string suppWbPath = @"D:\Земля 3 мини.xlsx";
            var suppPckg = new ExcelPackage(new FileInfo(suppWbPath));
            var suppWb = suppPckg.Workbook;

            var catWs = new CatalogWorksheet(suppWb.Worksheets["analytics"].ToDataTable());
            var oktmoDs = new DataSet();

            var dt1 = suppWb.Worksheets["нас.пункты РФ"].ToDataTable();
            var dt2 = suppWb.Worksheets["РегЦентры"].ToDataTable();
            oktmoDs.Tables.Add(dt1);
            oktmoDs.Tables.Add(dt2);

            var oktmoWs = new OKTMORepository(oktmoDs,"нас.пункты РФ");
            var subjWs = new SubjectSourceWorksheet(suppWb.Worksheets["Список источников по регионам"].ToDataTable());
            var vgtWs = new VGTWorksheet(suppWb.Worksheets["ВГТ"].ToDataTable());

            var supportWss = new SupportWorksheets(catWs,oktmoWs,subjWs,vgtWs);


            const string wbPAth = @"\\192.168.100.2\share\ДЛЯ______\Для Менеджеров БД\Денис\Инструменты\Обрабокта выгрузок\Обработать\14 год КН ЗУ\3 Земля\Result\А-К.xlsx";
            var pckg = new ExcelPackage(new FileInfo(wbPAth));
            var ws = pckg.Workbook.Worksheets[1];


            for (var i = 3; i < 100; i++)
            {
                var dataRow = new ExcelLocationRow(ws, i, XlTemplateWorkbookType.LandProperty, supportWss)
                {
                    DoDescription = false
                };
                dataRow.CheckRowForLocations();
            }

            pckg.SaveWithDialog();

            Assert.AreEqual(true,true);
        }
    }
}
