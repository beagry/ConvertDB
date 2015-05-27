using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using Converter.Template_workbooks;
using Converter.Template_workbooks.EFModels;
using Converter.Tools;
using ExcelRLibrary;
using Microsoft.Office.Interop.Excel;
using DataTable = System.Data.DataTable;
using TemplateWorkbook = Converter.Template_workbooks.TemplateWorkbook;

namespace Converter
{
    public class SourceWs
    {
        const int TakeFirstItemsQuantity = 300;

        private readonly List<int> checkedColumnsList;
        private readonly TemplateWorkbook templateWorkbook;

        private readonly Template_workbooks.EFModels.TemplateWorkbook wb;

        private readonly Dictionary<int, string> head; 
        private readonly DataTable wsTable;

        /// <summary>
        /// Key = номер столбца, который будет скопирован, Value = Название колонки Куда будет скопирован столбец
        /// </summary>
        private readonly Dictionary<int, string> columnsDictionary = new Dictionary<int, string>();

        public Dictionary<string, List<string>> ResultDictionary
        {
            get
            {
                return columnsDictionary
                    .Select(kp => new {ColumnCopy = head.First(hk => hk.Key == kp.Key).Value, ColumnPaste = kp.Value})
                    .GroupBy(obj => obj.ColumnPaste,o => o.ColumnCopy)
                    .ToDictionary(k => k.Key, v => v.ToList());
            }
        }

        public SourceWs(DataTable table, TemplateWorkbook templateWorkbook):this(templateWorkbook)
        {
            wsTable = table;
            head = wsTable.Columns.Cast<DataColumn>().ToDictionary(k => wsTable.Columns.IndexOf(k)+1, v => v.ColumnName);
        }

        public SourceWs(Worksheet worksheet, TemplateWorkbook templateWorkbook):this(templateWorkbook)
        {
            var sourceWorksheet = worksheet;

            wsTable = FillDataTable.GetDataTable(((Workbook) sourceWorksheet.Parent).FullName,
                sourceWorksheet.Name, TakeFirstItemsQuantity);
            head = worksheet.ReadHead();
        }

        public SourceWs(TemplateWorkbook workbook):this()
        {
            this.templateWorkbook = workbook;
        }

        public SourceWs(DataTable table, XlTemplateWorkbookType wbType)
            : this()
        {
            wsTable = table;
            head = wsTable.Columns.Cast<DataColumn>().ToDictionary(k => wsTable.Columns.IndexOf(k) + 1, v => v.ColumnName);

            var db = TemplateWbsRepository.Context;
            wb = db.TemplateWorkbooks.First(w => w.WorkbookType == wbType);
        }

        private SourceWs()
        {
            checkedColumnsList = new List<int>();            
        }
        
        public void CheckColumns()
        {
            //
            //Общие колонки
            //

            TryToFindTemplateColumnsFromDbData();
            return;
            TryToFindTemplateColumns();

//            GetSOURCE_LINK();
            GetDESCRIPTION();
            GetSUBJECT();
            GetREGION();
            GetNEAR_CITY();

            GetHEAT_SUPPLY();
            GetSYSTEM_ELECTRICITY();
            GetSYSTEM_GAS();
            GetSYSTEM_SEWERAGE();
            GetSYSTEM_WATER();

            GetPRICE();
            GetDateParsing();
            GetDATE_RESEARCH();
            GetAREA_LOT();
            GetSOURCE_DESC();
            GetOperationType();
            GetCONTACTS();

            //Уникальные поля Зем участков
            if (templateWorkbook is LandPropertyTemplateWorkbook)
            {
                GetRights();
                GetDIST_REG_CENTER();
                GetBuildings();
                GetLAND_CATEGORY();
                GetPERMITTED_USE();
                GetRELIEF();
                GetVEGETATION();                
                GetSeller();
            }
            //Уникальные поля Коммерции
            if (templateWorkbook is CommercePropertyTemplateWorkbook)
            {
                GetHEIGHT_FLOOR();
                GetMATERIAL_WALL();
                GetCONDITION();
                GetSECURITY();
                GetSEGMENT();
                GetBUILDING_TYPE();
                GetOBJECT_PURPOSE();
                GetFLOOR();
                GetFLOORITY();
                GetYEAR_BUILD();
                GetCLASS_TYPE();
            }

        }

        private void TryToFindTemplateColumnsFromDbData()
        {
            var db = TemplateWbsRepository.Context;
            var wbs = db.TemplateWorkbooks;
            var wb = wbs.First(t => t.WorkbookType == XlTemplateWorkbookType.LandProperty);
            var columns = wb.Columns;

            foreach (var column in columns)
            {
                var maskList = column.SearchCritetias.Select(s => s.Text).ToList();
                var columnCode = column.CodeName;
                GetColumnNumberByColumnName(columnCode, maskList);
            }
        }

        #region GetColumnMethods


        private void GetSUBJECT()
        {
            Dictionary<int, decimal> columnMatchDictionary = new Dictionary<int, decimal>();
            int result;
            const decimal percentIsOk = (decimal) 0.4;
            List<string> maskList = new List<string>(new[] {"СУБЪЕКТ_РОССИЙСКОЙ_ФЕДЕРАЦИИ","субъект","република", "область", "край"});
            if (GetColumnNumberByColumnName("SUBJECT", maskList)) return;
            //В каждой колонке поочередно
            for (int i = 0; i < wsTable.Columns.Count; i++)
            {
                if (checkedColumnsList.Contains(i + 1)) continue;
                if (wsTable.Columns[i].DataType != Type.GetType("System.String")) continue;

                decimal resultDecimal;
                //Берём все уникальные объекты
                List<string> uniqSubjsOfsourceWS = (from x in wsTable.AsEnumerable()
                    where !String.IsNullOrEmpty(x.Field<String>(i))
                    select x.Field<String>(i)).Distinct().ToList();

                int cup = (int) (percentIsOk*uniqSubjsOfsourceWS.Count());
                if (uniqSubjsOfsourceWS.Any(s => s.Contains("http") || s.Length > 100)) continue;
                int fitCellsQuantity =
                    uniqSubjsOfsourceWS.Count(
                        x => x.Contains(maskList[2]) || x.Contains(maskList[3]) || x.Contains(maskList[4]));

                if (fitCellsQuantity == 0) continue;
                //1.0 = 100% значений столбца
                resultDecimal = Decimal.Round((decimal) fitCellsQuantity/uniqSubjsOfsourceWS.Count(), 2);
                //Пишем результать совпадений в колонке
                columnMatchDictionary.Add(i, resultDecimal);
                if (resultDecimal >= percentIsOk) break;
            }
            if (columnMatchDictionary.Count == 0) return;
            result = columnMatchDictionary.Aggregate((l, r) => l.Value > r.Value ? l : r).Key + 1;

            columnsDictionary.Add(result, "SUBJECT");
            checkedColumnsList.Add(result);
        }

        private void GetREGION()
        {
            //Муниципальное образование 
            List<string> maskList = new List<string>(new[] {"МЕСТОПОЛОЖЕНИЕ","район", "город"});

            if (GetColumnNumberByColumnName("REGION",maskList)) return;
        }

        private void GetNEAR_CITY()
        {
            //Населенный пункт
            
            List<string> maskList = new List<string>(new[] { "населенн", "насел" });
            const string columnCode = "NEAR_CITY";
            if (GetColumnNumberByColumnName(columnCode, maskList)) return;

        }

        private void GetDESCRIPTION()
        {

            List<string> maskList = new List<string> { "ОПИСАНИЕ" };
            const string columnCode = "DESCRIPTION";
            if (GetColumnNumberByColumnName(columnCode, maskList)) return;

            if (columnsDictionary.ContainsValue(columnCode)) return;

            int c;
            int[] cups = {300, 150, 100};
            foreach (int cup in cups)
            {
                for (int i = 0; i < wsTable.Columns.Count; i++)
                {
                    if (checkedColumnsList.Contains(i + 1)) continue;
                    if (wsTable.Columns[i].DataType != Type.GetType("System.String")) continue;
                    //Просто находим столец, в котором очень много букв
                    c =
                        (wsTable.AsEnumerable().Where(x => !String.IsNullOrEmpty(x.Field<string>(i)))).Count(
                            x => x.Field<string>(i).Length > cup);
                    if (c == 0) continue;

                    //Нашли
                    columnsDictionary.Add(i + 1, columnCode);
                    checkedColumnsList.Add(i + 1);
                    return;
                }
            }
        }

        private void GetLAND_CATEGORY()
        {
            //"КАТЕГОРИЯ_ЗЕМЛИ"
            GetColumnNumberByColumnName("LAND_CATEGORY", new List<string> { "КАТЕГОРИЯ_ЗЕМЛИ", "категор", "земл" });
        }

        private void GetPERMITTED_USE()
        {
            List<string> maskList = new List<string>(new[] { "ВИД_РАЗРЕШЕННОГО_ИСПОЛЬЗОВАНИЯ", "вид р", "разрешен", "использ" });

            string columnCode = "PERMITTED_USE";
            if (GetColumnNumberByColumnName(columnCode, maskList)) return;
        }

        private void GetSOURCE_LINK()
        {

            List<string> maskList = new List<string>(new[] { "ССЫЛКА_НА_ИСТОЧНИК_ИНФОРМАЦИИ","ССЫЛКА" });

            string columnCode = String.Empty;
            columnCode = "SOURCE_LINK";

            if (columnCode == String.Empty) return;

            if (columnsDictionary.ContainsValue(columnCode)) return;

            if (GetColumnNumberByColumnName(columnCode, maskList)) return;


            for (int i = 0; i < wsTable.Columns.Count; i++)
            {
                if (checkedColumnsList.Contains(i + 1)) continue;
                if (wsTable.Columns[i].DataType != Type.GetType("System.String")) continue;
                if (
                    !wsTable.AsEnumerable()
                        .Where(x => !String.IsNullOrEmpty(x.Field<string>(i)))
                        .Any(x => x.Field<string>(i).Contains("http"))) continue;
                decimal percentSimilarity =
                    (decimal)
                        wsTable.AsEnumerable()
                            .Where(x2 => x2.Field<string>(i) != null)
                            .Count(x1 => x1.Field<String>(i).Contains("http"))/
                    (decimal) (wsTable.AsEnumerable().Where
                        (x2 => !String.IsNullOrEmpty(x2.Field<String>(i)))
                        )
                        .Count();
                if (percentSimilarity < 0.5M) continue;
                int result = i + 1;
                
                columnsDictionary.Add(result, columnCode);
                checkedColumnsList.Add(i + 1);
                return;
            }
        }

        private void GetDIST_REG_CENTER()
        {
            List<string> mask = new List<string> {"УДАЛЕННОСТЬ", "центр"};
            GetColumnNumberByColumnName("DIST_REG_CENTER", mask);
        }

        private void GetSYSTEM_GAS()
        {
            List<string> maskList = new List<string> {"ГАЗОСНАБЖЕНИЕ","газоснаб", "газос", "газ", "коммуникац", "коммуник", "комм"};
            if (GetColumnNumberByColumnName("SYSTEM_GAS", maskList)) return;
        }

        private void GetSYSTEM_WATER()
        {
            List<string> maskList = new List<string> {"ВОДОСНАБЖЕНИЕ", "водоснаб", "водос", "вод" };
            if (GetColumnNumberByColumnName("SYSTEM_WATER", maskList)) return;
        }

        private void GetSYSTEM_SEWERAGE()
        {
            //string mask = "канализ";
            List<string> maskList = new List<string> {"КАНАЛИЗАЦИЯ", "канализац", "канализ", "канал" };
            if (GetColumnNumberByColumnName("SYSTEM_SEWERAGE", maskList)) return;

        }

        private void GetSYSTEM_ELECTRICITY()
        {

            List<string> maskList = new List<string> {"ЭЛЕКТРОСНАБЖЕНИЕ", "электроснаб", "электрос", "электро","эле" };
            if (GetColumnNumberByColumnName("SYSTEM_ELECTRICITY", maskList)) return;
        }

        private void GetHEAT_SUPPLY()
        {

            List<string> maskList = new List<string> {"ТЕПЛОСНАБЖЕНИЕ", "теплоснаб", "тепл", "обогр", "отопл" };
            if (GetColumnNumberByColumnName("HEAT_SUPPLY", maskList)) return;
        }

        private void GetRELIEF()
        {
            List<string> maskList = new List<string> { "рельеф"};
            if (GetColumnNumberByColumnName("RELIEF", maskList)) return;
        }

        private void GetVEGETATION()
        {

            List<string> maskList = new List<string> { "растен" };
            if (GetColumnNumberByColumnName("VEGETATION", maskList)) return;
        }

        private void GetPRICE()
        {
            List<string> maskList = new List<string> { "СТОИМОСТЬ", "стоим", "цена", "продаж" };
            if (templateWorkbook is LandPropertyTemplateWorkbook)
                GetColumnNumberByColumnName("PRICE", maskList);
            else
                GetColumnNumberByColumnName("SALE_PRICE", maskList);
        }

        private void GetDATE_RESEARCH()
        {
            if (columnsDictionary.ContainsValue("DATE_RESEARCH")) return;
            if (GetColumnNumberByColumnName("DATE_RESEARCH", new List<string> { "ДАТА_РАЗМЕЩЕНИЯ_ИНФОРМАЦИИ", "ДАТА_РАЗМЕЩЕНИЯ", "дата" })) return;
        }

        private void GetAREA_LOT()
        {

            List<string> maskList = new List<string> {"ПЛОЩАДЬ УЧАСТКА", "ПЛОЩАДЬ_УЧАСТКА","ПЛОЩАДЬ", "площад", "площ" };
            if (GetColumnNumberByColumnName("AREA_LOT", maskList)) return;
        }

        private void GetCONTACTS()
        {
            GetColumnNumberByColumnName("CONTACTS", new List<string> { "ТЕЛЕФОН_ПРОДАВЦА" });
            GetColumnNumberByColumnName("CONTACTS", new List<string> { "КОНТАКТЫ" });
            GetColumnNumberByColumnName("CONTACTS", new List<string> { "ТЕЛЕФОН"});
        }

        private void GetOperationType()
        {
            List<string> masksList = new List<string> {"ВИД_СДЕЛКИ","ВИД СДЕЛКИ","продажа", "аренда"};

            if (GetColumnNumberByColumnName("OPERATION", masksList)) return;
        }       

        //BUG Ниже идут комерческие поля

        private void GetBALCONY()
        {
            GetColumnNumberByColumnName("BALCONY", new List<string>() { "ЛОДЖИЯ" });
            GetColumnNumberByColumnName("BALCONY", new List<string>() { "БАЛКОН" });
        }

        private void GetOBJECT_TYPE()
        {
            GetColumnNumberByColumnName("OBJECT_TYPE", new List<string>() { "ТИП_ОБЬЕКТА" });
        }

        private void GetAREA_TOTAL()
        {
            GetColumnNumberByColumnName("AREA_TOTAL", new List<string>() { "ПЛОЩАДЬ ОБЪЕКТА", "ПЛОЩАДЬ_ОБЪЕКТА", });
        }

        

        private void GetDateParsing()
        {
            List<string> maskList = new List<string> { "ДАТА_ПАРСИНГА" };
            GetColumnNumberByColumnName("DATE_PARSING", maskList);
        }

        private void GetFLOORITY()
        {
            List<string> maskList = new List<string> { "ЭТАЖНОСТЬ" };
            GetColumnNumberByColumnName("FLOOR_QNT_MIN", maskList);
        }

        private void GetCLASS_TYPE()
        {
            List<string> maskList = new List<string> { "ПОТРЕБИТЕЛЬСКИЙ_КЛАСС" };
            GetColumnNumberByColumnName("CLASS_TYPE", maskList);
        }

        private void GetYEAR_BUILD()
        {
            List<string> maskList = new List<string> { "ГОД_ПОСТРОЙКИ" };
            GetColumnNumberByColumnName("YEAR_BUILD", maskList);
        }

        private void GetFLOOR()
        {
            List<string> maskList = new List<string> { "ЭТАЖ" };
            GetColumnNumberByColumnName("FLOOR", maskList);
        }

        private void GetOBJECT_PURPOSE()
        {
            List<string> maskList = new List<string> { "НАЗНАЧЕНИЕ_ОБЪЕКТА" };
            GetColumnNumberByColumnName("OBJECT_PURPOSE", maskList);
        }

        private void GetBUILDING_TYPE()
        {
            List<string> maskList = new List<string> { "ТИП_ПОСТРОЙКИ" };
            GetColumnNumberByColumnName("BUILDING_TYPE", maskList);
        }

        private void GetSEGMENT()
        {
            List<string> maskList = new List<string> { "СЕГМЕНТ" };
            GetColumnNumberByColumnName("SEGMENT", maskList);
        }

        private void GetSECURITY()
        {
            List<string> maskList = new List<string> { "БЕЗОПАСНОСТЬ" };
            GetColumnNumberByColumnName("SECURITY", maskList);
        }

        private void GetCONDITION()
        {
            List<string> maskList = new List<string> { "СОСТОЯНИЕ" };
            GetColumnNumberByColumnName("CONDITION", maskList);
        }

        private void GetMATERIAL_WALL()
        {
            List<string> maskList = new List<string> { "МАТЕРИАЛ_СТЕН" };
            GetColumnNumberByColumnName("MATERIAL_WALL", maskList);
        }

        private void GetHEIGHT_FLOOR()
        {
            List<string> maskList = new List<string> { "ВЫСОТА_ПОТОЛКА" };
            GetColumnNumberByColumnName("HEIGHT_FLOOR", maskList);
        }

        private void GetRights()
        {
            List<string> maskList = new List<string> { "ВИД_ПРАВА","ВИД ПРАВА","права","прав"};
            GetColumnNumberByColumnName("LAW_NOW", maskList);
        }

        private void GetBuildings()
        {
            List<string> maskList = new List<string> { "СТРОЕНИЯ" };
            GetColumnNumberByColumnName("OBJECT", maskList);
        }

        private void GetSeller()
        {
            List<string> maskList = new List<string> { "КОМПАНИЯ" };
            GetColumnNumberByColumnName("SELLER", maskList);
        }

        private void GetSOURCE_DESC()
        {
            List<string> maskList = new List<string> { "ИСТОЧНИК_ИНФОРМАЦИИ","ИСТОЧНИК" };
            GetColumnNumberByColumnName("SOURCE_DESC", maskList);
        }

        #endregion

        /// <summary>
        /// Метод ищет колонки по названиями используя вшитые правила
        /// </summary>
        private void TryToFindTemplateColumns()
        {
            foreach (JustColumn templateColumn in templateWorkbook.TemplateColumns)
            {
                GetColumnNumberByColumnName(templateColumn.CodeName, new List<string>() { templateColumn.Description }, true);
            }
        }

        /// <summary>
        /// Метод находит колонку с полным или частичным совпадением в имени по пеерданному списку критериев поиска
        /// </summary>
        /// <param name="columnCode">Название колонки для записи результата</param>
        /// <param name="masks">Маски для сопоставления</param>
        /// <param name="fullSimilar">Обязательно полное совпадение</param>
        /// <returns></returns>
        private bool GetColumnNumberByColumnName(string columnCode, List<string> masks, bool fullSimilar = false)
        {

            masks.Add(columnCode);
            masks.Reverse();

            if (masks.Count == 0) return false;
            //Если мы уже находили такую колонку
            var c = 0;
            DataColumn cl;

            do //Поиск колонки с ПОЛНЫМ совпалением по одному из критериев маски поиска
            {
                cl = wsTable.Columns.Cast<DataColumn>().Where(x => !checkedColumnsList.Contains(x.Ordinal + 1)).
                            FirstOrDefault(x => string.Equals(x.ColumnName, masks[c], StringComparison.CurrentCultureIgnoreCase));
                c++;
            } while (cl == null & masks.Count - 1 >= c);


            //Поиск колонки с ЧАСТИЧНЫМ совпалением по одному из критериев маски поиска
            if (cl == null && !fullSimilar)
            {
                c = 0;
                do
                {
                    cl = wsTable.Columns.Cast<DataColumn>().Where(x => !checkedColumnsList.Contains(x.Ordinal + 1)).
                        FirstOrDefault(x => x.ColumnName.IndexOf(masks[c], StringComparison.OrdinalIgnoreCase) > -1);
                    c++;
                } while (cl == null & masks.Count - 1 >= c);
            }


            //ничего не нашли
            if (cl == null) return false;


            checkedColumnsList.Add(cl.Ordinal + 1);
            //В словарь результатов
            columnsDictionary.Add(cl.Ordinal + 1, columnCode);
            //Результат работы
            return true;
        }
    }
}
// ReSharper restore SuggestUseVarKeywordEvident