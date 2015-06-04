#define CheckHead
using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;
using Converter.Template_workbooks;
using Converter.Template_workbooks.EFModels;
using ExcelRLibrary;
using ExcelRLibrary.TemplateWorkbooks;
using Formater.SupportWorksheetsClasses;
using OfficeOpenXml;
using Excel = Microsoft.Office.Interop.Excel;


namespace Formater
{
    public delegate void VoidDelegate();

    public partial class DbToConvert
    {
        private const string noInfoString = "не указано";

        private ExcelPackage package;
        private ExcelWorksheet worksheet;
        private static int lastUsedRow = 0;
        private List<long> rowsToDelete;

        private CatalogWorksheet catalogWorksheet;
        private OKTMOWorksheet oktmo;
        private SubjectSourceWorksheet subjectSourceWorksheet;
        private VGTWorksheet vgtWorksheet;
        private MainForm MainForm;
        private ProgressBar progressBar;

        public int HeadSize { get; set; }

        public ExcelPackage ExcelPackage
        {
            get { return package; }
        }

        private readonly XlTemplateWorkbookType wbType;
        private readonly Dictionary<int, string> head; 
        private TemplateWbsContext db;


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

        private readonly byte subjColumn ;             
        private readonly byte regionColumn ;           
        private readonly byte settlementColumn ;       
        private readonly byte nearCityColumn ;         
        private readonly byte typeOfNearCityColumn ;   
        private readonly byte vgtColumn ;              
        private readonly byte streetColumn ;           
        private readonly byte typeOfStreetColumn ;     
        private readonly byte sourceLinkColumn ;       
        private readonly byte distToRegCenterColumn ;  
        private readonly byte distToNearCityColumn ;   
        private readonly byte inCityColumn ;           
        private readonly byte houseNumColumn ;         
        private readonly byte letterColumn ;           
        private readonly byte sntKpDnpColumn ;         
        private readonly byte additionalInfoColumn ;   
        private readonly byte buildColumn ;            


        public DbToConvert(MainForm mainForm, XlTemplateWorkbookType wbType) : this()
        {
            MainForm = mainForm;
            this.wbType = wbType;

            package = new ExcelPackage(new FileInfo(mainForm.WorkbookPath));
            worksheet = package.Workbook.Worksheets.First();
            lastUsedRow = worksheet.Dimension.Rows;
            head = worksheet.ReadHead();

            db = new TemplateWbsContext();
            var columns = db.TemplateWorkbooks.First(w => w.WorkbookType == wbType).Columns.ToList();
            subjColumn =               (byte)columns.First(c => c.CodeName.Equals("SUBJECT")).ColumnIndex;
            regionColumn =             (byte)columns.First(c => c.CodeName.Equals("REGION")).ColumnIndex;
            settlementColumn =         (byte)columns.First(c => c.CodeName.Equals("SETTLEMENT")).ColumnIndex;
            nearCityColumn =           (byte)columns.First(c => c.CodeName.Equals("NEAR_CITY")).ColumnIndex;
            typeOfNearCityColumn =     (byte)columns.First(c => c.CodeName.Equals("TERRITORY_TYPE")).ColumnIndex;
            vgtColumn =                (byte)columns.First(c => c.CodeName.Equals("VGT")).ColumnIndex;
            streetColumn =             (byte)columns.First(c => c.CodeName.Equals("STREET")).ColumnIndex;
            typeOfStreetColumn =       (byte)columns.First(c => c.CodeName.Equals("STREET_TYPE")).ColumnIndex;
            sourceLinkColumn =         (byte)columns.First(c => c.CodeName.Equals("URL_SALE")).ColumnIndex;
            distToRegCenterColumn =    (byte)columns.First(c => c.CodeName.Equals("DIST_REG_CENTER")).ColumnIndex;
            distToNearCityColumn =     (byte)columns.First(c => c.CodeName.Equals("DIST_NEAR_CITY")).ColumnIndex;
            inCityColumn =             (byte)columns.First(c => c.CodeName.Equals("IN_CITY")).ColumnIndex;
            houseNumColumn =           (byte)columns.First(c => c.CodeName.Equals("HOUSE_NUM")).ColumnIndex;
            letterColumn =             (byte)columns.First(c => c.CodeName.Equals("LETTER")).ColumnIndex;
            sntKpDnpColumn =           (byte)columns.First(c => c.CodeName.Equals("ASSOCIATIONS")).ColumnIndex;
            additionalInfoColumn =     (byte)columns.First(c => c.CodeName.Equals("ADDITIONAL")).ColumnIndex;
            buildColumn =              (byte)columns.First(c => c.CodeName.Equals("HOUSE_NUM")).ColumnIndex;
        }



        private DbToConvert()
        {
            HeadSize = 2;
            rowsToDelete = new List<long>();
        }

        public bool ColumnHeadIsOk()
        {
            var i = 1;
            var columns = db.TemplateWorkbooks.First(w => w.WorkbookType == wbType).Columns.ToList();
#if DEBUG
            foreach (var templateCode in columns.Select(c => c.CodeName))
            {
                if (worksheet.Cells[1, i].Value.ToString() != templateCode)
                {
                    MessageBox.Show(String.Format("Табличная шапка в листе {0} не соотвествует стандарту",
                        worksheet.Name));
                    return false;
                }
                i++;
            }
#endif
            var reader = new ExcelReader();
            oktmo = new OKTMOWorksheet(reader.ReadExcelFile(MainForm.OKTMOPath), MainForm.OKTMOWsName);

            subjectSourceWorksheet =
                new SubjectSourceWorksheet(
                    reader.ReadExcelFile(MainForm.SubjectLinkPath)
                        .Tables.Cast<DataTable>()
                        .First(t => t.TableName.Equals(MainForm.SubjectLinkWsName)));

            vgtWorksheet = new VGTWorksheet(reader.ReadExcelFile(MainForm.VGTPath)
                .Tables.Cast<DataTable>()
                .First(t => t.TableName.Equals(MainForm.VGTWsName)));

            catalogWorksheet = new CatalogWorksheet(reader.ReadExcelFile(MainForm.CatalogPath)
                .Tables.Cast<DataTable>()
                .First(t => t.TableName.Equals(MainForm.CatalogWsName)));

            lastUsedRow = worksheet.Dimension.End.Row;

            return true;
        }


        /// <summary>
        /// Общий метод, запускающий подметоды своего типа
        /// </summary>
        /// <returns></returns>
        public bool FormatWorksheet()
        {
            progressBar = MainForm.progressBar;
            if (worksheet == null ||
                oktmo == null ||
                catalogWorksheet == null) return false;


            FormatClassification();

            FormatCommunications();
            FormatAreaLot();
            FormatPrice();

            FormatOfferDeal();
            FormatOperation();
            FormatLandLaw();
            FormatSaleType();

            FormatLandCategory();

            FormatDate("DATE_RESEARCH");
            FormatDate("DATE_PARSING");
            FormatDate("DATE_IN_BASE");
            FormatBuildings();
            FormatLastUpdateDate();

            FormatSurface();
            FormatRoad();
            FormatRelief();

            FormatDistToRegCenter();

            return true;
        }

        #region Format Methods

        private void FormatRelief()
        {
            var columnIndex = GetColumnIndex("RELIEF");

            for (var i = HeadSize; i <= worksheet.Dimension.End.Row; i++)
            {
                var cell = worksheet.Cells[i, columnIndex];
                if (string.IsNullOrEmpty(cell.Value as string))
                {
                    cell.Value = noInfoString;
                    continue;
                }

                var val = cell.Value.ToString();

                if (Regex.IsMatch(val, "ровн", RegexOptions.IgnoreCase))
                {
                    cell.Value = "ровный";
                    continue;
                }
                if (Regex.IsMatch(val, @"не\sзнач", RegexOptions.IgnoreCase))
                {
                    cell.Value = "небольшой уклон";
                    continue;
                }
                if (Regex.IsMatch(val, "склон", RegexOptions.IgnoreCase))
                {
                    cell.Value = "склон";
                    continue;
                }
                if (Regex.IsMatch(val, "знач", RegexOptions.IgnoreCase))
                    cell.Value = "значительные перепады высот";
            }
        }

        private void FormatRoad()
        {
            var columnIndex = GetColumnIndex("ROAD");

            for (var i = HeadSize; i <= worksheet.Dimension.End.Row; i++)
            {
                var cell = worksheet.Cells[i, columnIndex];
                if (string.IsNullOrEmpty(cell.Value as string))
                {
                    cell.Value = noInfoString;
                    continue;
                }

                if (Regex.IsMatch(cell.Value.ToString(), "асф", RegexOptions.IgnoreCase))
                {
                    cell.Value = "асфальтовая дорога";
                    continue;
                }
                if (Regex.IsMatch(cell.Value.ToString(), "бетон", RegexOptions.IgnoreCase))
                {
                    cell.Value = "бетонка";
                    continue;
                }
                if (Regex.IsMatch(cell.Value.ToString(), "грун", RegexOptions.IgnoreCase))
                {
                    cell.Value = "грунтовая дорога";
                    continue;
                }
                if (Regex.IsMatch(cell.Value.ToString(), "грав", RegexOptions.IgnoreCase))
                    cell.Value = "гравийная дорога";
                else
                    cell.Value = noInfoString;
            }
        }

        private void FormatSurface()
        {
            var columnIndex = GetColumnIndex("SURFACE");

            for (var i = HeadSize; i <= worksheet.Dimension.End.Row; i++)
            {
                var cell = worksheet.Cells[i, columnIndex];
                if (string.IsNullOrEmpty(cell.Value as string))
                {
                    cell.Value = noInfoString;
                    continue;
                }

                if (Regex.IsMatch(cell.Value.ToString(), "асф", RegexOptions.IgnoreCase))
                {
                    cell.Value = "асфальт";
                    continue;
                }
                if (Regex.IsMatch(cell.Value.ToString(), "бетон", RegexOptions.IgnoreCase))
                {
                    cell.Value = "бетонные плиты";
                    continue;
                }
                if (Regex.IsMatch(cell.Value.ToString(), "грун", RegexOptions.IgnoreCase))
                    cell.Value = "грунт";
                else
                    cell.Value = noInfoString;
            }
        }

        private void FormatBuildings()
        {
            var columnIndex = GetColumnIndex("OBJECT");
            Regex regex = new Regex("дом", RegexOptions.IgnoreCase);
            for (var i = HeadSize; i <= worksheet.Dimension.End.Row; i++)
            {
                var cell = worksheet.Cells[i, columnIndex];
                if (string.IsNullOrEmpty(cell.Value as string))
                {
                    continue;
                }

                //Участки с домами удаляем
                if (regex.IsMatch(cell.Value.ToString()))
                {
                    rowsToDelete.Add(cell.Start.Row);
                }
                else
                    cell.Value = "да";
            }
        }

        private void FormatSaleType()
        {
            var columnIndex = GetColumnIndex("SALE_TYPE");
            var lawNowColumnIndex = GetColumnIndex("LAW_NOW");
            if (lawNowColumnIndex == 0) return;

            for (var i = HeadSize; i <= worksheet.Dimension.End.Row; i++)
            {
                var cell = worksheet.Cells[i, columnIndex];
                if (string.IsNullOrEmpty(cell.Value as string))
                {
                    continue;
                }
                cell.Value = worksheet.Cells[cell.Start.Row, lawNowColumnIndex].Value.ToString() == "аренда"
                    ? "переуступка прав аренды"
                    : "продажа";
            }
        }

        private void FormatOperation()
        {
            const string columnCode = "OPERATION";
            var columnIndex = GetColumnIndex(columnCode);
            var v = catalogWorksheet.GetContentByCode(columnCode);

            for (var i = HeadSize; i <= worksheet.Dimension.End.Row; i++)
            {
                var cell = worksheet.Cells[i, columnIndex];
                if (cell.Value == null) continue;
                if (v.Contains(cell.Value.ToString())) continue;
                cell.Value = "продажа";
            }
        }

        private void FormatOfferDeal()
        {
            var columnCode = "OFFER_DEAL";
            var columnIndex = GetColumnIndex(columnCode);
            var v = catalogWorksheet.GetContentByCode(columnCode);

            for (var i = HeadSize; i <= worksheet.Dimension.End.Row; i++)
            {
                var cell = worksheet.Cells[i, columnIndex];
                if (cell.Value == null) continue;
                if (v.Contains(cell.Value.ToString())) continue;

                cell.Value = "предложение";
            }
        }

        private void FormatLastUpdateDate()
        {
            const string parsingDateColumnName = "дата парсинга";
            const string parsingDateColumnName2 = "дата_парсинга";

            const string columnCode = "DATE_IN_BASE";


            var parsingColumn = GetColumnIndex("DATE_PARSING");

            var columnIndex = GetColumnIndex(columnCode);
            var dateRegex = new Regex("(сегодн|(поза)?вчер)");

            for (var i = HeadSize; i <= worksheet.Dimension.End.Row; i++)
            {
                var cell = worksheet.Cells[i, columnIndex];
                if (string.IsNullOrEmpty(cell.Value as string))
                {
                    cell.Value = noInfoString;
                }

                //Если есть дата
                var value = cell.Value.ToString();

                DateTime dt;
                if (cell.Value is DateTime || DateTime.TryParse(value,out dt)) continue;
                

                double u = 0;
                if (value is double|| double.TryParse(value,out u))
                    dt = DateTime.FromOADate(u);
                else
                    DateTime.TryParse((string) value, out dt);

                Match match;
                //Когда не удалось конвертиорвать в дату
                if (dt < new DateTime(2000, 01, 01))
                {
                    Regex regex = new Regex(@"\d\d\.\d\d\.\d{2,4}");
                    match = regex.Match(value);
                    if (match.Success)
                    {
                        value = match.Value;
                        DateTime.TryParse(value, out dt);
                    }
                }

                //Есть всё хорошо и мы нашли нормальную дату
                if (dt > new DateTime(2000, 01, 01))
                {
                    cell.Value = dt;
                    cell.Style.Numberformat.Format = "dd.mm.yyyy";
                    continue;
                }

                //Если не нашли колонку с датой парсинга
                if (parsingColumn == 0) continue;

                //Если нет даты
                match = dateRegex.Match(value);
                if (!match.Success) continue;

                var days = 0;
                if (match.Value == "позавчер")
                    days = -2;
                if (match.Value == "сегодн")
                    days = 0;
                if (match.Value == "вчер")
                    days = -1;


                if (worksheet.Cells[cell.Start.Row, parsingColumn].Value == null) continue;

                value = worksheet.Cells[cell.Start.Row, parsingColumn].Value.ToString();
                DateTime.TryParse((string) value, out dt);

                if (dt < new DateTime(2000, 01, 01)) continue;
                dt = dt.AddDays(days);
                cell.Value = dt;
                cell.Style.Numberformat.Format = "dd.mm.yyyy";
            }
            progressBar.Invoke(new VoidDelegate(() => progressBar.Value += 10)); //Инкрементируем прогрессбар
        }

        private void FormatDate(string columnCode)
        {
            var columnIndex = GetColumnIndex(columnCode);
            for (var i = HeadSize; i <= worksheet.Dimension.End.Row; i++)
            {
                var cell = worksheet.Cells[i, columnIndex];
                if (string.IsNullOrEmpty(cell.Value as string))
                {
                    continue;
                }


                if (cell.Style.Numberformat.Format == "dd.mm.yyyy") continue;

                if (cell.Value is DateTime)
                {
                    cell.Style.Numberformat.Format = "dd.mm.yyyy";
                    continue;
                }
                object value = cell.Value;
                DateTime dt;
                if (value == null) continue;
                if (value is double)
                {
                    dt = DateTime.FromOADate((double) value);
                }
                else
                {
                    DateTime.TryParse((string) value, out dt);
                }
                if (dt < new DateTime(2000, 01, 01))
                {
                    Regex regex = new Regex(@"\d\d\.\d\d\.\d{2,4}");
                    Match match = regex.Match(cell.Value.ToString());
                    value = match.Value;
                    DateTime.TryParse((string) value, out dt);
                    if (dt < new DateTime(2000, 01, 01))
                    {
                        dt = TryPasreDate(cell.Value.ToString());
                        if (dt == DateTime.MinValue) continue;
                    }
                    ;
                }
                cell.Value = dt < new DateTime(2000, 01, 01) ? (dynamic) string.Empty : dt;
                cell.Style.Numberformat.Format = "dd.mm.yyyy";
            }
        }

        private DateTime TryPasreDate(string text)
        {
            var dict = new Dictionary<string, string>()
            {
                {"янв", "01"},
                {"февр", "02"},
                {"март", "03"},
                {"апр", "04"},
                {"ма(й|я)", "05"},
                {"июн", "06"},
                {"июл", "07"},
                {"авг", "08"},
                {"сент", "09"},
                {"окт", "10"},
                {"нояб", "11"},
                {"дек", "12"},
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
            var columnIndex = GetColumnIndex("LAND_CATEGORY");

            for (var i = HeadSize; i <= worksheet.Dimension.End.Row; i++)
            {
                var cell = worksheet.Cells[i, columnIndex];
                if (string.IsNullOrEmpty(cell.Value as string))
                {
                    cell.Value = noInfoString;
                    continue;
                }

                if (Regex.IsMatch(cell.Value.ToString(), "сельхо", RegexOptions.IgnoreCase) ||
                    Regex.IsMatch(cell.Value.ToString(), "с.х", RegexOptions.IgnoreCase))
                {
                    cell.Value = "Земли сельскохозяйственного назначения";
                    continue;
                }
                if (Regex.IsMatch(cell.Value.ToString(), "пром", RegexOptions.IgnoreCase))
                {
                    cell.Value = "Земли промышленности и иного назначения";
                    continue;
                }
                if (Regex.IsMatch(cell.Value.ToString(), "селен", RegexOptions.IgnoreCase))
                {
                    cell.Value = "Земли населенных пунктов";
                    continue;
                }

                //Последняя проверка и запись дефолтного значения
                var regex = new Regex(@"(охран|лесн|водн|запас)", RegexOptions.IgnoreCase);
                if (regex.IsMatch(cell.Value.ToString()))
                    rowsToDelete.Add(cell.Start.Row);
                else
                {
                    cell.Value = "Земли населенных пунктов";
                }
            }
        }

        private void FormatLandLaw()
        {
            const string code = "LAW_NOW";
            var columnIndex = GetColumnIndex(code);
            var v = catalogWorksheet.GetContentByCode(code);
            var rentalPeriodColumnIndex = GetColumnIndex("RENTAL_PERIOD");

            for (var i = HeadSize; i <= worksheet.Dimension.End.Row; i++)
            {
                var cell = worksheet.Cells[i, columnIndex];

                if (string.IsNullOrEmpty(cell.Value as string))
                {
                    cell.Value = "собственность";
                    continue;
                }

                var value = cell.Value.ToString();

                if (v.Contains(value)) continue;

                if (Regex.IsMatch(value, "аренд", RegexOptions.IgnoreCase))
                {
                    Match match = new Regex(@"\d+").Match(value);
                    if (match.Success && rentalPeriodColumnIndex != 0)
                        worksheet.Cells[cell.Start.Row, rentalPeriodColumnIndex].Value =
                            String.Format("на {0} лет",
                                match.Value);
                    cell.Value = "аренда";
                }
                else
                    cell.Value = "собственность";


            }
        }


        private void FormatCommunications()
        {
            var firstColumnIndex = GetColumnIndex("SYSTEM_GAS");

            var lastColumnIndex = GetColumnIndex("HEAT_SUPPLY");

            var usingRange = worksheet.Cells[2, firstColumnIndex,lastUsedRow, lastColumnIndex];

            var columnsCodeList = new List<string>
            {
                "SYSTEM_GAS",
                "SYSTEM_WATER",
                "SYSTEM_SEWERAGE",
                "SYSTEM_ELECTRICITY",
                "HEAT_SUPPLY",
            };

            var columnIndex = firstColumnIndex;
            var v = catalogWorksheet.GetContentByCode(columnsCodeList[0]);

            //Проверяем первый столбец(Газ) на предмет информации для соседних столбцов
            for (var i = HeadSize; i <= worksheet.Dimension.End.Row; i++)
            {
                var cell = worksheet.Cells[i, columnIndex];
                if (string.IsNullOrEmpty(cell.Value as string))
                {
                    cell.Value = noInfoString;
                    continue;
                }

                if (v.Contains(cell.Value.ToString())) continue;

                //IDEA а что если в первой ячейке есть "Вод" а в столбце "водоснабжение" есть "родниковая вода"
                //Прочие элементарные вараинты для Газа
                if (Regex.IsMatch(cell.Value.ToString(), @"Нет", RegexOptions.IgnoreCase) &&
                    cell.Value.ToString().Length < 6)
                {
                    cell.Value = "отсутствует, возможность подключения неизвестна";
                    continue;
                }
                if (Regex.IsMatch(cell.Value.ToString(), @"есть", RegexOptions.IgnoreCase) &&
                    cell.Value.ToString().Length < 6)
                {
                    cell.Value = "есть, но не указано какое";
                    continue;
                }


                if (Regex.IsMatch(cell.Value.ToString(), "электр", RegexOptions.IgnoreCase))
                    worksheet.Cells[
                        cell.Start.Row, GetColumnIndex("SYSTEM_ELECTRICITY")
                        ].Value =
                        "есть, выделенная мощность неизвестна";
                if (Regex.IsMatch(cell.Value.ToString(), "вод", RegexOptions.IgnoreCase))
                    worksheet.Cells[cell.Start.Row, GetColumnIndex("SYSTEM_WATER")]
                        .Value =
                        "есть, но не указано какое";
                //просто пишем поверх
                if (Regex.IsMatch(cell.Value.ToString(), "скваж", RegexOptions.IgnoreCase))
                    worksheet.Cells[cell.Start.Row, GetColumnIndex("SYSTEM_WATER")]
                        .Value =
                        "скважина";
                if (Regex.IsMatch(cell.Value.ToString(), "канализ", RegexOptions.IgnoreCase))
                    worksheet.Cells[cell.Start.Row, GetColumnIndex("SYSTEM_SEWERAGE")]
                        .Value =
                        "есть, но не указано какое";
                if (Regex.IsMatch(cell.Value.ToString(), "отопл", RegexOptions.IgnoreCase))
                    worksheet.Cells[cell.Start.Row, GetColumnIndex("HEAT_SUPPLY")]
                        .Value =
                        "есть, но не указано какое";
                //В самом конце можем перезаписать ячейку
                cell.Value = Regex.IsMatch(cell.Value.ToString(), "газ", RegexOptions.IgnoreCase)
                    ? "есть, но не указано какое"
                    : noInfoString;
            }

            //Проходимся по остальным столбцам (Вода, электр-во, канализация, отопление)
            //Варианты :Есть, Нет, Пусто

            //По всем столбцам
            for (var n = 2; n <= columnsCodeList.Count; n++)
            {
                var columnName = columnsCodeList[n - 1];
                //Значения по справочнику
                v = catalogWorksheet.GetContentByCode(columnName);

                columnIndex = GetColumnIndex(columnName);
                //Далее по всем ячейкам в столбце
                for (var i = HeadSize; i <= worksheet.Dimension.End.Row; i++)
                {
                    var cell = worksheet.Cells[i, columnIndex];
                    if (string.IsNullOrEmpty(cell.Value as string))
                    {
                        cell.Value = noInfoString;
                        continue;
                    }

                    if (v.Contains(cell.Value.ToString())) continue;
                    //когда "10 вкт"
                    if (Regex.IsMatch(cell.Value.ToString(), @"\bквт\b", RegexOptions.IgnoreCase))
                        continue;

                    if (Regex.IsMatch(cell.Value.ToString(), @"нет", RegexOptions.IgnoreCase) &&
                        cell.Value.ToString().Length < 6)
                        cell.Value = "отсутствует, возможность подключения неизвестна";
                    if (Regex.IsMatch(cell.Value.ToString(), @"есть", RegexOptions.IgnoreCase) &&
                        cell.Value.ToString().Length < 6)
                    {
                        cell.Value = n == 4 ? "есть, выделенная мощность неизвестна" : "есть, но не указано какое";
                    }
                }
            }
            progressBar.Invoke(new VoidDelegate(() => progressBar.Value += 20)); //Инкрементируем прогрессбар
        }

        private void FormatPrice()
        {
            const string columnCode = "PRICE";
            var columnIndex = GetColumnIndex(columnCode);

            var numericRegex = new Regex(@"(\d|\s|\.|\,)+");
            var multiplierRegex = new Regex(@"(г(ект)?а|сот|(/)?(кв\\s*\\.?\\s*м\\b|м2|м\\s*\\.\\s*кв\b|м\b))",
                RegexOptions.IgnoreCase);

            for (var i = HeadSize; i <= worksheet.Dimension.End.Row; i++)
            {
                var cell = worksheet.Cells[i, columnIndex];
                if (cell.Value == null) return;

                double u;
                //Цифровые ячейки пропускаем
                if (cell.Value is double || double.TryParse(cell.Value.ToString(), out u))
                {
                    if (cell.Value.ToString() == 0.ToString(CultureInfo.InvariantCulture))
                        cell.Value = string.Empty;
                    continue;
                }

                var x = 1; //Множитель для млн и тысячи в числа
                var multReg = new Regex(@"(\b(млн|тыс))", RegexOptions.IgnoreCase);

                Match convertMatch;
                try
                {
                    convertMatch = multReg.Match(cell.Value.ToString());
                }
                catch (Exception)
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

                var cellValue = cell.Value.ToString();

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
                        worksheet.Cells[cell.Start.Row, GetColumnIndex("AREA_LOT")].Value is
                            double)
                    {
                        multiplier =
                            (double) worksheet.Cells[cell.Start.Row, GetColumnIndex("AREA_LOT")].Value;
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
                            cell.Value = String.Empty;
                            continue;
                        }

                        //За метр квадратный
                        //ТО есть есть общая площадь (что не всегда) и есть за м.кв.
                        var pricePerUnitCell =
                            worksheet.Cells[cell.Start.Row, GetColumnIndex("PRICE_FOR_UNIT")];

                        double s2;

                        //Убераем пробемы, заменяем точку на запятую и конвертирует в double
                        Double.TryParse(match.Value.Trim().Replace(" ", String.Empty).Replace(".", ","), out s2);

                        pricePerUnitCell.Value = s2*y*x;
                        pricePerUnitCell.Style.Numberformat.Format = "#";
                    }
                }

                //Проверка на 85.000
                var val = Regex.IsMatch(cell.Value.ToString(), @"\d+\.\d{3,}")
                    ? cell.Value.ToString().Replace(".", String.Empty)
                    : cellValue;

                match = numericRegex.Match(val);
                if (!match.Success)
                {
                    cell.Value = String.Empty;
                    continue;
                }

                double s;

                //Убераем пробемы, заменяем точку на запятую и конвертирует в double
                Double.TryParse(match.Value.Trim().Replace(" ", String.Empty).Replace(".", ","), out s);

                cell.Value = s*y*multiplier*x;
                cell.Style.Numberformat.Format = "#";
            }
        }

        private void FormatAreaLot()
        {
            const string columnCode = "AREA_LOT";

            var columnIndex = GetColumnIndex(columnCode);

            //10 000,89 руб / 9 000.80 рубсотк
            Regex numericRegex = new Regex(@"(\d|\s|\.|\,)+");
            Regex multiplieRegex = new Regex(@"(га|сот)", RegexOptions.IgnoreCase);

            for (var i = HeadSize; i <= worksheet.Dimension.End.Row; i++)
            {
                var cell = worksheet.Cells[i, columnIndex];
                if (cell.Value == null) return;

                double u;
                //Цифровые ячейки пропускаем
                if (cell.Value is double || double.TryParse(cell.Value.ToString(), out u))
                {
                    if (cell.Value.ToString() == 0.ToString(CultureInfo.InvariantCulture))
                        cell.Value = string.Empty;
                    continue;
                }

                //когда в ячейке площадь дома, мы не обрабатываем участки с домом
                if (Regex.IsMatch(cell.Value.ToString(), "дом", RegexOptions.IgnoreCase))
                {
                    rowsToDelete.Add(cell.Start.Row);
                    continue;
                }

                //Дефолтный мнжитель
                var y = 1;
                Match match;

                //При наличие модификатора Га или Сот
                //Берём первый встречный и меняем множитель
                if (multiplieRegex.IsMatch(cell.Value.ToString()))
                {
                    match = multiplieRegex.Match(cell.Value.ToString());
                    y = match.Value.ToLower() == "га" ? 10000 : 100;
                }

                //Вычленяем цифры с запятыми, точками и пробелом (которые сразу и режем)
                match = null;
                match = numericRegex.Match(cell.Value.ToString());
                if (!match.Success)
                {
                    cell.Value = String.Empty;
                    continue;
                }
                double s;
                Double.TryParse(match.Value.Trim().Replace(" ", String.Empty).Replace(".", ","), out s);

                cell.Value = s*y;
                cell.Style.Numberformat.Format = @"#";
            }
        }

        #endregion

        internal int GetColumnIndex(string columnCode)
        {
            var col = head.First(p => p.Value.Equals(columnCode));
            return col.Key;
        }


        /// Метод бэкапит данные, подвергшиеся замещению
//        private void ReserveColumns()
//        {
//            string wsName = "Reserve";
//
//            if (ColumnsToReserve == null || ColumnsToReserve.Count == 0) return;
//
//            Excel.Workbook workbook = worksheet.Parent as Excel.Workbook;
//            if (workbook == null) return;
//
//            Excel.Worksheet reserveWorksheet =
//                workbook.Worksheets.Cast<Excel.Worksheet>().FirstOrDefault(ws => ws.Name == wsName) ??
//                workbook.Worksheets.Add(Type.Missing, worksheet, Type.Missing, Type.Missing);
//
//            reserveWorksheet.Name = wsName;
//
//
//            foreach (string columnName in ColumnsToReserve)
//            {
//                string name = columnName;
//
//                int srcColumn = LandPropertyTemplateWorkbook.GetColumnByCode(name);
//                int lastColumn = srcColumn;
//
//                Excel.Range trgtRange = reserveWorksheet.Columns[lastColumn] as Excel.Range;
//                Excel.Range srcRange = worksheet.Columns[srcColumn] as Excel.Range;
//                if (trgtRange != null && srcRange != null)
//                    trgtRange.Value = srcRange.Value;
//            }
//        }
    }
}
