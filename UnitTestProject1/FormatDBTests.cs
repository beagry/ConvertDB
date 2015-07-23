using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Controls;
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
        public void StartFormat()
        {
            IFormatDbParams viewModel = new FormatDbViewModel();
//            viewModel.Path = @"B:\Managers\Денис\Инструменты\Обрабокта выгрузок\Обработать\14 год КН ЗУ\3 Земля\Result\С-Я.xlsx";
//            viewModel.Path = @"M:\Managers\Денис\Инструменты\Обрабокта выгрузок\Этап 2\Топ.xlsx";
            viewModel.Path = @"M:\Managers\Денис\Инструменты\Обрабокта выгрузок\Этап 2\Zy_irr1411_typed.xlsx";

            const string suppWbPath = @"D:\Земля 3 мини.xlsx";
            viewModel.CatalogSupportWorkbook.Path = suppWbPath;
            viewModel.CatalogSupportWorkbook.SelectedWorksheet = "analytics";
            viewModel.VgtCatalogSupportWorkbook.Path = suppWbPath;
            viewModel.VgtCatalogSupportWorkbook.SelectedWorksheet = "ВГТ";
            viewModel.OktmoSupportWorkbook.Path = suppWbPath;
            viewModel.OktmoSupportWorkbook.SelectedWorksheet = "нас.пункты РФ";
            viewModel.SubjectSourceSupportWorkbook.Path = suppWbPath;
            viewModel.SubjectSourceSupportWorkbook.SelectedWorksheet = "Список источников по регионам";
            
            viewModel.DoDescription = false;

            var convert = new DbToConvert(viewModel)
            {
                ColumnsToReserve = new List<string> { "SUBJECT", "REGION", "NEAR_CITY", "SYSTEM_GAS", "SYSTEM_WATER", "SYSTEM_SEWERAGE", "SYSTEM_ELECTRICITY" },
                DoDescription = false
            };

            var checkHeadResult = convert.ColumnHeadIsOk();
            if (!checkHeadResult) Assert.AreEqual(false,true);

            //Запусть обработки в новом потоке
            convert.FormatWorksheet();

            convert.SaveResult();
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

//            var reader = new ExcelReader();
//            var ds = reader.ReadExcelFile(@"B:\Managers\Шапки и справочники\Справочники\STREET_new.xlsx");
//            var kladr = new KladrRepository(ds.Tables.Cast<DataTable>().First(t => t.TableName.Equals("STREET")));
            var kladr = new KladrRepository();
            var oktmoWs = new OKTMORepository(oktmoDs,"нас.пункты РФ",kladr);
            var subjWs = new SubjectSourceWorksheet(suppWb.Worksheets["Список источников по регионам"].ToDataTable());
            var vgtWs = new VGTWorksheet(suppWb.Worksheets["ВГТ"].ToDataTable());
            

            var supportWss = new SupportWorksheets(catWs,oktmoWs,subjWs,vgtWs,kladr);


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
