#define CheckHead
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;
using Formater.SupportWorksheetsClasses;
using Excel = Microsoft.Office.Interop.Excel;


namespace Formater
{
    public delegate void VoidDelegate();
    public partial class DbToConvert
    {
        private Excel.Worksheet worksheet;
        private Excel.Application xlApplication;
        private static long lastUsedRow = 0;
        private List<long> rowsToDelete;

        private CatalogWorksheet catalogWorksheet;
        private OKTMOWorksheet oktmo;
        private SubjectSourceWorksheet subjectSourceWorksheet ;
        private VGTWorksheet vgtWorksheet;
        private MainForm MainForm;
        private ProgressBar progressBar;


#if !DEBUG
        public string WorkbookPath { get; set; }
        public string OKTMOPath { get; set; }
        public string OKTMOWsName { get; set; }
        public string CatalogPath { get; set; }
        public string CatalogWsName { get; set; }
        public string SubjectLinkPath { get; set; }
        public string SubjectLinkWsName { get; set; }
        public string VGTPath { get; set; }
        public string VGTWsName { get; set; }
#else
        private const string CatalogWsName = "analytics";
        private const string OKTMOWorksheetname = "нас.пункты РФ";
        private const string sourceSubjWsName = "Список источников по регионам";
        private const string vgtWsName = "ВГТ";
#endif

        public List<string> ColumnsToReserve { get; set; }

        public Excel.Application XlApplication { get { return xlApplication; } }
        
        private readonly byte subjColumn = (byte) LandPropertyTemplateWorkbook.GetColumnByCode("SUBJECT");
        private readonly byte regionColumn = (byte) LandPropertyTemplateWorkbook.GetColumnByCode("REGION");
        readonly byte settlementColumn = (byte) LandPropertyTemplateWorkbook.GetColumnByCode("SETTLEMENT");
        readonly byte nearCityColumn = (byte) LandPropertyTemplateWorkbook.GetColumnByCode("NEAR_CITY");
        readonly byte typeOfNearCityColumn = (byte) LandPropertyTemplateWorkbook.GetColumnByCode("TERRITORY_TYPE");
        readonly byte vgtColumn = (byte) LandPropertyTemplateWorkbook.GetColumnByCode("VGT");
        readonly byte streetColumn = (byte) LandPropertyTemplateWorkbook.GetColumnByCode("STREET");
        readonly byte typeOfStreetColumn = (byte) LandPropertyTemplateWorkbook.GetColumnByCode("STREET_TYPE");
        readonly byte sourceLinkColumn = (byte) LandPropertyTemplateWorkbook.GetColumnByCode("URL_SALE");
        readonly byte distToRegCenterColumn = (byte) LandPropertyTemplateWorkbook.GetColumnByCode("DIST_REG_CENTER");
        readonly byte distToNearCityColumn = (byte)LandPropertyTemplateWorkbook.GetColumnByCode("DIST_NEAR_CITY");
        readonly byte inCityColumn = (byte) LandPropertyTemplateWorkbook.GetColumnByCode("IN_CITY");
        readonly byte houseNumColumn = (byte) LandPropertyTemplateWorkbook.GetColumnByCode("HOUSE_NUM");
        readonly byte letterColumn = (byte) LandPropertyTemplateWorkbook.GetColumnByCode("LETTER");
        readonly byte sntKpDnpColumn = (byte)LandPropertyTemplateWorkbook.GetColumnByCode("ASSOCIATIONS");
        readonly byte additionalInfoColumn = (byte)LandPropertyTemplateWorkbook.GetColumnByCode("ADDITIONAL");
        readonly byte buildColumn = (byte) LandPropertyTemplateWorkbook.GetColumnByCode("HOUSE_NUM");
        
        public DbToConvert(MainForm mainForm)
        {
            MainForm = mainForm;

            xlApplication = LandPropertyTemplateWorkbook.GetExcelApplication();
            rowsToDelete = new List<long>();
        }

        

        ~DbToConvert()
        {
//            if (worksheet.Parent != null)
//            {
//                if (!worksheet.Parent.Equals())
//            }
//            if (xlApplication != null)
//            {
//                xlApplication.ScreenUpdating = true;
//                xlApplication.Visible = true;
//            }
        }

        public bool SaveResult()
        {
            XlApplication.Visible = true;
            xlApplication.ScreenUpdating = true;
            using (SaveFileDialog fileDialog = new SaveFileDialog())
            {
                fileDialog.Filter = @"Excel File|*.xlsx";
                fileDialog.DefaultExt = @"*.xlsx";
                fileDialog.Title = @"Выберите место для сохранения";
                fileDialog.FileName = "Обработанная выгрузка";

                if (fileDialog.ShowDialog() != DialogResult.OK) return false;

                Excel.Workbook workbook = worksheet.Parent;
                try
                {
                    xlApplication.EnableEvents = false;
                    workbook.SaveAs(fileDialog.FileName, Excel.XlFileFormat.xlOpenXMLWorkbook);
                    xlApplication.EnableEvents = true;
//                    workbook.Close();
                }
                catch (COMException e)
                {
                    xlApplication.EnableEvents = true;
                    MessageBox.Show(e.Message,
                                    String.Format("Ошибка при сохранении. \nОбработанная книга висит с названием {0}. \nОсторожней, не закройте случайно :)", workbook.Name),MessageBoxButtons.OK,MessageBoxIcon.Error);
                    return false;
                }
#if !DEBUG
                oktmo.CloseWorkbook();
                subjectSourceWorksheet.CloseWorkbook();
                catalogWorksheet.CloseWorkbook();
                vgtWorksheet.CloseWorkbook();
#endif 
                    //var workbook1 = workbook.Parent as Excel.Workbook;
                //if (workbook1 != null) workbook1.Close();

                return true;
            }
        }

        private void DeteleRows()
        {
            if (rowsToDelete.Count == 0) return;

            var sortedRows = rowsToDelete.OrderByDescending(x => x);

            foreach (var row in sortedRows.Select(row => worksheet.Rows[row]).Cast<Excel.Range>().Select(x => x.Row))
            {
                ((Excel.Range) worksheet.Rows[row]).Interior.Color = Color.DimGray;
//                worksheet.Rows[row].EntireRow.Delete(Excel.XlDeleteShiftDirection.xlShiftUp);
            }
        }

        public bool ColumnHeadIsOk()
        {
#if !DEBUG
            worksheet = xlApplication.Workbooks.Open(MainForm.WorkbookPath,false,true).Worksheets["База"];
#else
            xlApplication.EnableEvents = false;
            worksheet = xlApplication.Workbooks.Open(@"D:\Земля 3.xlsx",false,true).Worksheets[1];
            xlApplication.EnableEvents = true;
#endif

#if !DEBUG
            var i = 1;
            foreach (var templateCode in LandPropertyTemplateWorkbook.TemplateColumns.Select(x => x.Code))
            {
                if (worksheet.Cells[1, i].Value2.ToString() !=templateCode)
                {
                    MessageBox.Show(String.Format("Табличная шапка в листе {0} книги {1} не соотвествует стандарту",worksheet.Name, worksheet.Parent.Name));
                    return false;
                }
                i++;
            }

#endif
            //IF OK INITIAL CLASS
#if DEBUG
            //worksheet = xlApplication.Workbooks["база сборка2.xlsm"].Worksheets["база"];
            catalogWorksheet = new CatalogWorksheet((worksheet.Parent as Excel.Workbook).Worksheets[CatalogWsName]);
            oktmo = new OKTMOWorksheet((worksheet.Parent as Excel.Workbook).Worksheets[OKTMOWorksheetname]);
            subjectSourceWorksheet = new SubjectSourceWorksheet((worksheet.Parent as Excel.Workbook).Worksheets[sourceSubjWsName]);
            vgtWorksheet = new VGTWorksheet((worksheet.Parent as Excel.Workbook).Worksheets[vgtWsName]);
#else
            oktmo = new OKTMOWorksheet(xlApplication.Workbooks.Open(MainForm.OKTMOPath,false,true).Worksheets[MainForm.OKTMOWsName]);
            subjectSourceWorksheet =
                new SubjectSourceWorksheet(xlApplication.Workbooks.Open(MainForm.SubjectLinkPath, false, true).Worksheets[MainForm.SubjectLinkWsName]);
            vgtWorksheet = new VGTWorksheet(xlApplication.Workbooks.Open(MainForm.VGTPath, false, true).Worksheets[MainForm.VGTWsName]);
            catalogWorksheet = new CatalogWorksheet(xlApplication.Workbooks.Open(MainForm.CatalogPath, false, true).Worksheets[MainForm.CatalogWsName]);
#endif
            try
            {
                worksheet.ShowAllData();
            }
            catch (COMException e)
            {
                if (e.HResult != -2146827284) throw;
            }
            xlApplication.Selection.EntireRow.AutoFit();
            //For reset UsedRange
            var t1 = worksheet.UsedRange.Rows.Count;
            lastUsedRow = worksheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row;


            return true;
        }

        
        /// <summary>
        /// Общий метод, запускающий подметоды своего типа
        /// </summary>
        /// <returns></returns>
        public async Task<bool> FormatWorksheet() //async Task<bool>
        {
            progressBar = MainForm.progressBar;
            //statusBar.BeginInvoke(new DoVoid(() => statusBar.Value = 30));
            if (xlApplication.Interactive == false) return false; //Edit Mode is on?
            if (worksheet == null ||
                oktmo == null ||
                catalogWorksheet == null) return false;


//            FormatClassification();
#if (!DEBUG)

//            FormatCommunications();
            FormatAreaLot();
            FormatPrice();

//            FormatOfferDeal();
//            FormatOperation();
//            FormatLandLaw();
//            FormatSaleType();
//
//            FormatLandCategory();

            FormatDate("DATE_RESEARCH");
            FormatDate("DATE_PARSING");
            FormatDate("DATE_IN_BASE");
//            FormatBuildings();
//            FormatLastUpdateDate();
//
//            FormatSurface();
//            FormatRoad();
//            FormatRelief();
//
//            var range = SetColumnRange("DESCRIPTION");
//            range.Cells.WrapText = false;
//
#endif
//
//            FormatDistToRegCenter();
//            DeteleRows();
            return true;
        }

        #region Format Methods
        

        private void FormatRelief()
        {
            var columnRange = SetColumnRange("RELIEF");
            columnRange.Replace(String.Empty, "не указано");

            foreach (var cell in columnRange.Cast<Excel.Range>().Where(cell => cell.Value2 != null).Where(cell => cell.Value2.ToString() == "не указано"))
            {
                if (Regex.IsMatch(cell.Value2.ToString(), "ровн", RegexOptions.IgnoreCase))
                {
                    cell.Value2 = "ровный";
                    continue;
                }
                if (Regex.IsMatch(cell.Value2.ToString(), @"не\sзнач", RegexOptions.IgnoreCase))
                {
                    cell.Value2 = "небольшой уклон";
                    continue;
                }
                if (Regex.IsMatch(cell.Value2.ToString(), "склон", RegexOptions.IgnoreCase))
                {
                    cell.Value2 = "склон";
                    continue;
                }
                if (Regex.IsMatch(cell.Value2.ToString(), "знач", RegexOptions.IgnoreCase))
                    cell.Value2 = "значительные перепады высот";
            }
        }

        private void FormatRoad()
        {
            var columnRange = SetColumnRange("ROAD");
            columnRange.Replace(String.Empty, "не указано");

            foreach (var cell in columnRange.Cast<Excel.Range>().Where(cell => cell.Value2 != null).Where(cell => cell.Value2.ToString() == "не указано"))
            {
                if (Regex.IsMatch(cell.Value2.ToString(), "асф", RegexOptions.IgnoreCase))
                {
                    cell.Value2 = "асфальтовая дорога";
                    continue;
                }
                if (Regex.IsMatch(cell.Value2.ToString(), "бетон", RegexOptions.IgnoreCase))
                {
                    cell.Value2 = "бетонка";
                    continue;
                }
                if (Regex.IsMatch(cell.Value2.ToString(), "грун", RegexOptions.IgnoreCase))
                {
                    cell.Value2 = "грунтовая дорога";
                    continue;
                }
                if (Regex.IsMatch(cell.Value2.ToString(), "грав", RegexOptions.IgnoreCase))
                    cell.Value2 = "гравийная дорога";
                else
                    cell.Value2 = "не указано";
            }
        }

        private void FormatSurface()
        {
            var columnRange = SetColumnRange("SURFACE");
            columnRange.Replace(String.Empty, "не указано");

            foreach (var cell in columnRange.Cast<Excel.Range>().Where(cell => cell.Value2 != null).Where(cell => cell.Value2.ToString() == "не указано"))
            {
                if (Regex.IsMatch(cell.Value2.ToString(), "асф", RegexOptions.IgnoreCase))
                {
                    cell.Value2 = "асфальт";
                    continue;
                }
                if (Regex.IsMatch(cell.Value2.ToString(), "бетон", RegexOptions.IgnoreCase))
                {
                    cell.Value2 = "бетонные плиты";
                    continue;
                }
                if (Regex.IsMatch(cell.Value2.ToString(), "грун", RegexOptions.IgnoreCase))
                    cell.Value2 = "грунт";
                else
                    cell.Value2 = "не указано";

            }
        }
        private void FormatBuildings()
        {
            var columnRange = SetColumnRange("OBJECT");
            Regex regex = new Regex("дом",RegexOptions.IgnoreCase);
            foreach (Excel.Range cell in columnRange.Cast<Excel.Range>().Where(cell => cell.Value2 != null))
            {
                //Участки с домами удаляем
                if (regex.IsMatch(cell.Value2.ToString()))
                {
                    rowsToDelete.Add(cell.Row);
                }
                else
                    cell.Value2 = "да";
            }
        }

        private void FormatSaleType()
        {
            var columnRange = SetColumnRange("SALE_TYPE");
            var lawNowColumnIndex = LandPropertyTemplateWorkbook.GetColumnByCode("LAW_NOW");
            if (lawNowColumnIndex == 0) return;

            foreach (Excel.Range cell in columnRange)
            {
                cell.Value2 = worksheet.Cells[cell.Row, lawNowColumnIndex].Value2.ToString() =="аренда" ? "переуступка прав аренды" : "продажа";
            }
        }

        private void FormatOperation()
        {
            const string columnCode = "OPERATION";
            var columnRange = SetColumnRange(columnCode);
            var v = catalogWorksheet.GetContentByCode(columnCode);

            foreach (Excel.Range cell in columnRange.Cast<Excel.Range>().Where(cell => cell.Value2 == null || !v.Contains(cell.Value2.ToString())))
            {
//                if (cell.Value2 == null || !v.Contains(cell.Value2.ToString())) 
                    cell.Value2 = "продажа";
            }
        }

        private void FormatOfferDeal()
        {
            var columnCode = "OFFER_DEAL";
            var columnRange = SetColumnRange(columnCode);
            var v = catalogWorksheet.GetContentByCode(columnCode);

            foreach (Excel.Range cell in columnRange.Cast<Excel.Range>().Where(cell => cell.Value2 == null || !v.Contains(cell.Value2.ToString())))
            {
//                if (cell.Value2 == null) 
                    cell.Value2 = "предложение";
            }
        }

        private void FormatLastUpdateDate()
        {

            const string parsingDateColumnName = "дата парсинга";
            const string parsingDateColumnName2 = "дата_парсинга";
            const string columnCode = "DATE_IN_BASE";
            Excel.Range parsingCell = null;

            try
            {
                parsingCell = (worksheet.UsedRange.Rows[1] as Excel.Range).Find(parsingDateColumnName) ??
                                              (worksheet.UsedRange.Rows[1] as Excel.Range).Find(parsingDateColumnName2);
            }
            catch (Exception)
            {

                MessageBox.Show(String.Format("Ошибка при проверке колонки \"{0}\".\n Колонка пропущена.", columnCode));
            }

            var parsingColumn = parsingCell == null? 0: parsingCell.Column;
            parsingCell = null;
            var columnRange = SetColumnRange(columnCode);
            DateTime dt;
            Regex dateRegex = new Regex("(сегодн|(поза)?вчер)");

            foreach (var cell in columnRange.Cast<Excel.Range>().Where(x=> x.Value != null))
            {
                if (cell.Value is DateTime) continue;

                //Если есть дата
                object value = cell.Value2;
                Match match;

                if (value is double)
                    dt = DateTime.FromOADate((double)value);
                else
                    DateTime.TryParse((string)value, out dt);

                //Когда не удалось конвертиорвать в дату
                if (dt < new DateTime(2000, 01, 01))
                {
                    Regex regex = new Regex(@"\d\d\.\d\d\.\d{2,4}");
                    match = regex.Match(cell.Value2);
                    if (match.Success)
                    {
                        value = match.Value;
                        DateTime.TryParse((string)value, out dt);
                    }
                }

                //Есть всё хорошо и мы нашли нормальную дату
                if (dt > new DateTime(2000, 01, 01))
                {
                    cell.Value2 = dt;
                    cell.NumberFormat = "dd.mm.yyyy";
                    continue;
                }

                //Если не нашли колонку с датой парсинга
                if (parsingColumn == 0) continue;

                //Если нет даты
                match = dateRegex.Match(cell.Value2);
                if (!match.Success) continue;

                var days = 0;
                if (match.Value == "позавчер")
                    days = -2;
                if (match.Value == "сегодн")
                    days = 0;
                if (match.Value == "вчер")
                    days = -1;
                
                
                
                if(worksheet.Cells[cell.Row, parsingColumn].Value == null) continue;

                value = worksheet.Cells[cell.Row, parsingColumn].Value2.ToString();
                DateTime.TryParse((string)value, out dt);

                if (dt < new DateTime(2000, 01, 01)) continue;
                dt = dt.AddDays(days);
                cell.Value = dt;
                cell.NumberFormat = "dd.mm.yyyy";
            }
            progressBar.Invoke(new VoidDelegate(() => progressBar.Value += 10));//Инкрементируем прогрессбар
        }

        private void FormatDate(string columnCode)
        {
            var columnRange = SetColumnRange(columnCode);
            foreach (var cell in columnRange.Cast<Excel.Range>().Where(cell => cell.Value2 != null).Where(x => x.Value2.ToString() != ""))
            {
                if (cell.NumberFormat == "dd.mm.yyyy") continue;
                if (cell.Value is DateTime)
                {
                    cell.NumberFormat = "dd.mm.yyyy";
                    continue;
                }
                object value = cell.Value2;
                DateTime dt;
                if (value == null) continue;
                if (value is double)
                {
                    dt = DateTime.FromOADate((double)value);
                }
                else
                {
                    DateTime.TryParse((string)value, out dt);
                }
                if (dt < new DateTime(2000, 01, 01))
                {
                    Regex regex = new Regex(@"\d\d\.\d\d\.\d{2,4}");
                    Match match =  regex.Match(cell.Value2.ToString());
                    value = match.Value;
                    DateTime.TryParse((string) value, out dt);
                    if (dt < new DateTime(2000, 01, 01))
                    {
                        dt = TryPasreDate(cell.Value2.ToString());
                        if (dt == DateTime.MinValue) continue;
                    };
                }
                cell.Value2 =dt < new DateTime(2000,01,01)? (dynamic) String.Empty :dt;
                cell.NumberFormat = "dd.mm.yyyy";
            }
        }

        private DateTime TryPasreDate(string text)
        {
            var dict = new Dictionary<string, string>()
            {
                {"янв","01"},
                {"февр","02"},
                {"март","03"},
                {"апр","04"},
                {"ма(й|я)","05"},
                {"июн","06"},
                {"июл","07"},
                {"авг","08"},
                {"сент","09"},
                {"окт","10"},
                {"нояб","11"},
                {"дек","12"},
            };

            foreach (var keyPair in dict)
            {
                var reg = new Regex("\\b" + keyPair.Key + "\\w*");
                if (!reg.IsMatch(text)) continue;

                text = reg.Replace(text, "." + keyPair.Value + ".");
                break;
            }

            text = text.Replace(" ", "");

            var dateReg = new Regex(@"\d\d\.\d\d\.\d{2,4}");

            var match = dateReg.Match(text);
            if (!match.Success) return DateTime.MinValue;

            return DateTime.Parse(match.Value);
        }

        private void FormatLandCategory()
        {
            var columnRange = SetColumnRange("LAND_CATEGORY");

            foreach (var cell in columnRange.Cast<Excel.Range>())
            {
                if (cell.Value2 == null || cell.Value2.ToString() == "")
                    cell.Value2 = "не указано";
                else
                {
                    if (Regex.IsMatch(cell.Value2.ToString(), "сельхо", RegexOptions.IgnoreCase) ||
                        Regex.IsMatch(cell.Value2.ToString(), "с.х", RegexOptions.IgnoreCase))
                    {
                        cell.Value2 = "Земли сельскохозяйственного назначения";
                        continue;
                    }
                    if (Regex.IsMatch(cell.Value2.ToString(), "пром", RegexOptions.IgnoreCase))
                    {
                        cell.Value2 = "Земли промышленности и иного назначения";
                        continue;
                    }
                    if (Regex.IsMatch(cell.Value2.ToString(), "селен", RegexOptions.IgnoreCase))
                    {
                        cell.Value2 = "Земли населенных пунктов";
                        continue;
                    }

                    //Последняя проверка и запись дефолтного значения
                    var regex = new Regex(@"(охран|лесн|водн|запас)",RegexOptions.IgnoreCase);
                    if (regex.IsMatch(cell.Value2.ToString()))
                        rowsToDelete.Add(cell.Row);
                    else
                    {
                        cell.Value2 = "Земли населенных пунктов"; 
                    }
                }
            }
        }

        private void FormatLandLaw()
        {
            const string code = "LAW_NOW";
            var columnRange = SetColumnRange(code);
            var v = catalogWorksheet.GetContentByCode(code);
            var rentalPeriodColumnIndex = LandPropertyTemplateWorkbook.GetColumnByCode("RENTAL_PERIOD");

            foreach (var cell in columnRange.Cast<Excel.Range>().Where(cell =>cell.Value2 == null || !v.Contains(cell.Value2.ToString())))
            {
                if (cell.Value2 == null)
                    cell.Value2 = "собственность";
                else
                {
                    if (Regex.IsMatch(cell.Value2.ToString(), "аренд", RegexOptions.IgnoreCase))
                    {
                        Match match = new Regex(@"\d+").Match(cell.Value2.ToString());
                        if (match.Success && rentalPeriodColumnIndex != 0)
                            worksheet.Cells[cell.Row, rentalPeriodColumnIndex].Value2 = String.Format("на {0} лет",
                                match.Value);
                        cell.Value2 = "аренда";
                    }
                    //else if (Regex.IsMatch(cell.Value2, "собст", RegexOptions.IgnoreCase))
                    //    cell.Value2 = "собственность";
                    //else if (Regex.IsMatch(cell.Value2, "частн", RegexOptions.IgnoreCase))
                    //    cell.Value2 = "собственность";
                    else
                        cell.Value2 = "собственность";
                }
            }
        }

        private void FormatCommunications()
        {
            var firstColumnIndex = LandPropertyTemplateWorkbook.TemplateColumns.First(x => x.Code == "SYSTEM_GAS").Index;
            var lastColumnIndex = LandPropertyTemplateWorkbook.TemplateColumns.First(x => x.Code == "HEAT_SUPPLY").Index;
            var usingRange =worksheet.Range[worksheet.Cells[2, firstColumnIndex], 
                                                     worksheet.Cells[lastUsedRow, lastColumnIndex]];
            var columnsCodeList = new List<string>
            {
                "SYSTEM_GAS",
                "SYSTEM_WATER",
                "SYSTEM_SEWERAGE",
                "SYSTEM_ELECTRICITY",
                "HEAT_SUPPLY",
            };
            Excel.Range columnRange = usingRange.Columns[1].Cells;
            var v = catalogWorksheet.GetContentByCode(columnsCodeList[0]);
            //Проверяем первый столбец(Газ) на предмет информации для соседних столбцов
            foreach (Excel.Range cell in columnRange)
            {
                if (cell.Value2 == null || cell.Value2.ToString() == "")
                    cell.Value2 = "не указано";
                else
                {
                    if (v.Contains(cell.Value2.ToString())) continue;
                    
                    //IDEA а что если в первой ячейке есть "Вод" а в столбце "водоснабжение" есть "родниковая вода"
                    //Прочие элементарные вараинты для Газа
                    if (Regex.IsMatch(cell.Value2.ToString(), @"Нет", RegexOptions.IgnoreCase) &&
                        cell.Value2.ToString().Length < 6)
                    {
                        cell.Value2 = "отсутствует, возможность подключения неизвестна";
                        continue;
                    }
                    if (Regex.IsMatch(cell.Value2.ToString(), @"есть", RegexOptions.IgnoreCase) &&
                        cell.Value2.ToString().Length < 6)
                    {
                        cell.Value2 = "есть, но не указано какое";
                        continue;
                    }


                    if (Regex.IsMatch(cell.Value2.ToString(), "электр", RegexOptions.IgnoreCase))
                        worksheet.Cells[cell.Row, LandPropertyTemplateWorkbook.GetColumnByCode("SYSTEM_ELECTRICITY")
                            ].Value2 =
                            "есть, выделенная мощность неизвестна";
                    if (Regex.IsMatch(cell.Value2.ToString(), "вод", RegexOptions.IgnoreCase))
                        worksheet.Cells[cell.Row, LandPropertyTemplateWorkbook.GetColumnByCode("SYSTEM_WATER")]
                            .Value2 =
                            "есть, но не указано какое";
                    //просто пишем поверх
                    if (Regex.IsMatch(cell.Value2.ToString(), "скваж", RegexOptions.IgnoreCase))
                        worksheet.Cells[cell.Row, LandPropertyTemplateWorkbook.GetColumnByCode("SYSTEM_WATER")]
                            .Value2 =
                            "скважина";
                    if (Regex.IsMatch(cell.Value2.ToString(), "канализ", RegexOptions.IgnoreCase))
                        worksheet.Cells[cell.Row, LandPropertyTemplateWorkbook.GetColumnByCode("SYSTEM_SEWERAGE")]
                            .Value2 =
                            "есть, но не указано какое";
                    if (Regex.IsMatch(cell.Value2.ToString(), "отопл", RegexOptions.IgnoreCase))
                        worksheet.Cells[cell.Row, LandPropertyTemplateWorkbook.GetColumnByCode("HEAT_SUPPLY")]
                            .Value2 =
                            "есть, но не указано какое";
                    //В самом конце можем перезаписать ячейку
                    cell.Value2 = Regex.IsMatch(cell.Value2.ToString(), "газ", RegexOptions.IgnoreCase)
                        ? "есть, но не указано какое"
                        : "не указано";
                }
            }

            //Проходимся по остальным столбцам (Вода, электр-во, канализация, отопление)
            //Варианты :Есть, Нет, Пусто
            
            //По всем столбцам
            for (var n = 2; n <= usingRange.Columns.Count; n++)
            {
                //Значения по справочнику
                v = catalogWorksheet.GetContentByCode(columnsCodeList[n-1]);
                
                columnRange = usingRange.Columns[n].Cells;
                //Далее по всем ячейкам в столбце
                foreach (Excel.Range cell in columnRange)
                {
                    if (cell.Value2 == null || cell.Value2.ToString() =="")
                        cell.Value2 = "не указано";
                    else
                    {
                        if (v.Contains(cell.Value2.ToString())) continue;
                        //когда "10 вкт"
                        if (Regex.IsMatch(cell.Value2.ToString(), @"\bквт\b", RegexOptions.IgnoreCase))
                           continue;

                        if (Regex.IsMatch(cell.Value2.ToString(), @"нет", RegexOptions.IgnoreCase) &&
                                cell.Value2.ToString().Length < 6)
                            cell.Value2 = "отсутствует, возможность подключения неизвестна";
                        if (Regex.IsMatch(cell.Value2.ToString(), @"есть", RegexOptions.IgnoreCase) &&
                                cell.Value2.ToString().Length < 6)
                        {
                            cell.Value2 = n == 4 ? "есть, выделенная мощность неизвестна" : "есть, но не указано какое";
                        }
                    }
                }
            }
            progressBar.Invoke(new VoidDelegate(() => progressBar.Value += 20));//Инкрементируем прогрессбар
        }

        private void FormatPrice()
        {
            const string columnCode = "PRICE";
            var columnRange = SetColumnRange(columnCode);

            var numericRegex = new Regex(@"(\d|\s|\.|\,)+");
            var multiplierRegex = new Regex(@"(г(ект)?а|сот|(/)?(кв\\s*\\.?\\s*м\\b|м2|м\\s*\\.\\s*кв\b|м\b))", RegexOptions.IgnoreCase);           

            foreach (var cell in columnRange.Cast<Excel.Range>().Where(x =>x.Value2 != null))
            {
                //Цифровые ячейки пропускаем
                if (cell.Value is double)
                {
                    if (cell.Value2.ToString() == 0.ToString(CultureInfo.InvariantCulture)) 
                        cell.Value2 = String.Empty;
                    continue;
                }

                var x = 1; //Множитель для млн и тысячи в числа
                var multReg = new Regex(@"(\b(млн|тыс))", RegexOptions.IgnoreCase);

                Match convertMatch;
                try
                {
                    convertMatch = multReg.Match(cell.Value2.ToString());
                }
                catch (Exception e)
                {
                    continue;
                }

                if (convertMatch.Success)
                {
                    if (convertMatch.Value.ToLower() == "млн")
                        x = 1000000;
                    else if (convertMatch.Value.ToLower() == "тыс")
                        x = 1000;
                }

                var y = 1; //Множитель для сотка-гектар
                double multiplier = 1; //множитель для "за сотку-метр-гектар". Содержит площадь земельного участка

                var cellValue = cell.Value2.ToString();

                //Проверяем строку на предмет "100 000рублей за сотку-метр-гектар"
                Match match = multiplierRegex.Match(cellValue);
                if (match.Success)
                {   
                    switch (match.Value.ToLower())
                    {
                        case "гекта":
                        case "га":
                            y = 1/10000;
                            break;
                        case "сот":
                            y = 1/100;
                            break;
                    }

                    if (
                        worksheet.Cells[cell.Row, LandPropertyTemplateWorkbook.GetColumnByCode("AREA_LOT")].Value is
                            double)
                    {
                        multiplier =
                            worksheet.Cells[cell.Row, LandPropertyTemplateWorkbook.GetColumnByCode("AREA_LOT")].Value;
                    }
                    else
                    {
                        //Проверка на 85.000
                        var val2 = Regex.IsMatch(cell.Value.ToString(), @"\d+\.\d{3,}")
                            ? cell.Value.ToString().Replace(".", String.Empty)
                            : cellValue;

                        match = numericRegex.Match(val2);
                        if (!match.Success)
                        {
                            cell.Value2 = String.Empty;
                            continue;
                        }

                        //За метр квадратный
                        //ТО есть есть общая площадь (что не всегда) и есть за м.кв.
                        var pricePerUnitCell =
                                worksheet.Cells[cell.Row, LandPropertyTemplateWorkbook.GetColumnByCode("PRICE_FOR_UNIT")];

                        double s2;

                        //Убераем пробемы, заменяем точку на запятую и конвертирует в double
                        Double.TryParse(match.Value.Trim().Replace(" ", String.Empty).Replace(".", ","), out s2);

                        pricePerUnitCell.Value = s2 * y * x;
                        pricePerUnitCell.NumberFormat = "#";

                    }

                }

                match = null;
                //Проверка на 85.000
                var val = Regex.IsMatch(cell.Value.ToString(), @"\d+\.\d{3,}")
                    ? cell.Value.ToString().Replace(".", String.Empty)
                    : cellValue;

                match = numericRegex.Match(val);
                if (!match.Success)
                {
                    cell.Value2 = String.Empty;
                    continue;
                }

                double s;
                
                //Убераем пробемы, заменяем точку на запятую и конвертирует в double
                Double.TryParse(match.Value.Trim().Replace(" ", String.Empty).Replace(".", ","), out s);

                cell.Value = s*y*multiplier * x;
                cell.NumberFormat = "#";
            }
        }

        private void FormatAreaLot()
        {
            const string columnCode = "AREA_LOT";          

            var columnRange = SetColumnRange(columnCode);

            //10 000,89 руб / 9 000.80 рубсотк
            Regex numericRegex = new Regex(@"(\d|\s|\.|\,)+");
            Regex multiplieRegex = new Regex(@"(га|сот)",RegexOptions.IgnoreCase);

            foreach (var cell in columnRange.Cast<Excel.Range>().Where(x => x.Value2 != null))
            {

                if (cell.Value is double)
                {
                    if (cell.Value == 0) cell.Value2 = string.Empty;
                    continue;
                }

                //когда в ячейке площадь дома, мы не обрабатываем участки с домом
                if (Regex.IsMatch(cell.Value2.ToString(), "дом", RegexOptions.IgnoreCase))
                {
                    rowsToDelete.Add(cell.Row);
                    continue;
                }

                //Дефолтный мнжитель
                var y = 1;
                Match match;

                //При наличие модификатора Га или Сот
                //Берём первый встречный и меняем множитель
                if (multiplieRegex.IsMatch(cell.Value2.ToString()))
                {
                    match = multiplieRegex.Match(cell.Value2.ToString());
                    y = match.Value.ToLower() == "га" ? 10000 : 100;
                }

                //Вычленяем цифры с запятыми, точками и пробелом (которые сразу и режем)
                match = null;
                match = numericRegex.Match(cell.Value2.ToString());
                if (!match.Success)
                {
                    cell.Value2 = String.Empty;
                    continue;
                }
                //var s = match.Value.Trim().Replace(" ",String.Empty);
                double s;
                Double.TryParse(match.Value.Trim().Replace(" ", String.Empty).Replace(".",","), out s);

                cell.Value = s * y;
                cell.NumberFormat = @"#";
            }
        }

        #endregion

        private Excel.Range SetColumnRange(string columnCode)
        {
            if (LandPropertyTemplateWorkbook.TemplateColumns.FirstOrDefault(x => x.Code == columnCode) == null)
                return null;

            var columnIndex = (byte)LandPropertyTemplateWorkbook.TemplateColumns.First(x => x.Code == columnCode).Index;
            var columnRange =
                worksheet.Range[worksheet.Cells[2, columnIndex], worksheet.Cells[lastUsedRow, columnIndex]].Cells;

            return columnRange;
        }


        private void ReserveColumns()
        {
            string wsName = "Reserve";

            if (ColumnsToReserve == null || ColumnsToReserve.Count == 0) return;
            
            Excel.Workbook workbook = worksheet.Parent as Excel.Workbook;
            if (workbook==null)return;

            Excel.Worksheet reserveWorksheet =
                workbook.Worksheets.Cast<Excel.Worksheet>().FirstOrDefault(ws => ws.Name == wsName) ??
                workbook.Worksheets.Add(Type.Missing, worksheet, Type.Missing, Type.Missing);

            reserveWorksheet.Name = wsName;


            foreach (string columnName in ColumnsToReserve)
            {
                string name = columnName;
// ReSharper disable once AssignNullToNotNullAttribute
                //Excel.Range row1 = worksheet.Rows[1] as Excel.Range;
                //if (row1 == null) continue;

                //Excel.Range firstOrDefault = row1.Cells.Cast<Excel.Range>().FirstOrDefault(cell => cell.Value2 != null && cell.Value2.ToString() == name);

                //if (firstOrDefault == null) continue;

                int srcColumn = LandPropertyTemplateWorkbook.GetColumnByCode(name); //firstOrDefault.Column;
                int lastColumn = srcColumn;

                Excel.Range trgtRange = reserveWorksheet.Columns[lastColumn] as Excel.Range;
                Excel.Range srcRange = worksheet.Columns[srcColumn] as Excel.Range;
                if (trgtRange != null && srcRange != null)
                    trgtRange.Value2 = srcRange.Value2;
            }
        }
    }
}
