using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text.RegularExpressions;
using System.Windows.Media.TextFormatting;
using Converter.Template_workbooks;
using Converter.Tools;
using ExcelRLibrary;
using DataTable = System.Data.DataTable;
using Excel = Microsoft.Office.Interop.Excel;


namespace Converter
{
    public class SourceWs
    {
        const int TakeFirstItemsQuantity = 300;

        private readonly List<int> checkedColumnsList;
        private readonly TemplateWorkbook templateWorkbook;

        private readonly Dictionary<int, string> head; 
        private readonly DataTable wsTable;

        /// <summary>
        /// Key = номер столбца, который будет скопирован, Value = Название колонки Куда будет скопирован столбец
        /// </summary>
        private readonly Dictionary<int, string> columnsDictionary = new Dictionary<int, string>();

//        private readonly List<JustColumn> sourceColumns;

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
            checkedColumnsList = new List<int>();

            wsTable = table;
            head = wsTable.Columns.Cast<DataColumn>().ToDictionary(k => wsTable.Columns.IndexOf(k)+1, v => v.ColumnName);
        }

        public SourceWs(Excel.Worksheet worksheet, TemplateWorkbook templateWorkbook):this(templateWorkbook)
        {
            var sourceWorksheet = worksheet;
            checkedColumnsList = new List<int>();

            wsTable = FillDataTable.GetDataTable(((Excel.Workbook) sourceWorksheet.Parent).FullName,
                sourceWorksheet.Name, TakeFirstItemsQuantity);
            head = worksheet.ReadHead();
        }

        public SourceWs(TemplateWorkbook workbook)
        {
            this.templateWorkbook = workbook;
        }
        
        public void CheckColumns()
        {
            //
            //Общие колонки
            //

            TryToFindTemplateColumns();

            GetSOURCE_LINK();
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

            if (templateWorkbook is CountryLiveAreaTemplateWorkbook ||
                templateWorkbook is CityLivaAreaTemplateWorkbook)
            {
                GetAREA_TOTAL();
                GetOBJECT_TYPE();
                GetBALCONY();
            }

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

//            JustColumn firstOrDefault = sourceColumns.FirstOrDefault(x => x.Index == result - 1);
//            if (firstOrDefault != null)
//                firstOrDefault.CodeName = "SUBJECT";

            columnsDictionary.Add(result, "SUBJECT");
            checkedColumnsList.Add(result);
        }

        private void GetREGION()
        {
            //Муниципальное образование 

            //Попытка 1
            //Столбец с наибольшим кол-вом (область)|(край)
            Dictionary<int, decimal> columnMatchDictionary = new Dictionary<int, decimal>();
            JustColumn firstOrDefault;
            //
            const decimal itemsIsOk = 0.55M;
            List<string> maskList = new List<string>(new[] {"МЕСТОПОЛОЖЕНИЕ","район", "город"});

            if (GetColumnNumberByColumnName("REGION",maskList)) return;

            //В каждой колонке поочередно
//            for (int i = 0; i < wsTable.Columns.Count; i++)
//            {
//
//                if (checkedColumnsList.Contains(i + 1)) continue;
//                if (wsTable.Columns[i].DataType != Type.GetType("System.String")) continue;
//                //Рабочий массив строк
//                List<string> workItemsList = (from x in wsTable.AsEnumerable()
//                    where !String.IsNullOrEmpty(x.Field<String>(i))
//                    select x.Field<String>(i)).Distinct().ToList();
//                int cup = (int) (workItemsList.Count*itemsIsOk);
//                //Ситауции когда сразу идём к следующей колонке
//                if (workItemsList.Any(s => s.Contains("http") || s.Length > 100)) continue;
//
//                //Проеряем на наличие слов поиска
//                int fitCellsQuantity = workItemsList.Count(x => x.Contains(maskList[2]) || x.Contains(maskList[1]));
//                //int cup = (int)(percentIsOk * uniqSubjsOfsourceWS.Count());
//
//                //Столбец не содержит ни одного их подходящих слов
//                if (fitCellsQuantity == 0) continue;
//
//
//                //Console.WriteLine(uniqSubjsOfsourceWS.Count());
//                //Пишем результать совпадений в колонке
//                columnMatchDictionary.Add(i, fitCellsQuantity);
//                if (fitCellsQuantity > cup) break;
//            }
//
//            if (columnMatchDictionary.Count == 0) return;
//            int result = columnMatchDictionary.Aggregate((l, r) => l.Value > r.Value ? l : r).Key + 1;
//            columnsDictionary.Add(result, "SETTLEMENT");
//            checkedColumnsList.Add(result);
//            firstOrDefault = sourceColumns.FirstOrDefault(x => x.Index == result - 1);
//            if (firstOrDefault != null)
//                firstOrDefault.CodeName = "SETTLEMENT";
        }

        private void GetNEAR_CITY()
        {
            //Населенный пункт
            
            List<string> maskList = new List<string>(new[] { "населенн", "насел" });
            const string columnCode = "NEAR_CITY";
            if (GetColumnNumberByColumnName(columnCode, maskList)) return;


//
//            const decimal percentIsOk = (decimal)0.6;
//            Dictionary<int, decimal> columnMatchDictionary = new Dictionary<int, decimal>();
//            int result;
            ////Массив для поиска
            //List<string> oktmoCurrColumnList = (from cell in oktmoTable.AsEnumerable()
            //    where !String.IsNullOrEmpty(cell.Field<string>("Название населенного пункта"))
            //    select cell.Field<string>("Название населенного пункта")).Distinct().ToList();
            ////В каждой колонке поочередно
            //for (int i = 0; i < wsTable.Columns.Count; i++)
            //{
            //    if (checkedColumnsList.Contains(i + 1)) continue;
            //    if (wsTable.Columns[i].DataType != Type.GetType("System.String")) continue;


            //    //Вce объекты
            //    List<string> totalSubjsOfsourceWS = (from row in wsTable.AsEnumerable()
            //        where !String.IsNullOrEmpty(row.Field<string>(i))
            //        select row.Field<String>(i)).ToList();
            //    if (totalSubjsOfsourceWS.Any(s => s.Contains("http") || s.Length > 100)) continue;

            //    //Получаем процент схожести между 
            //    decimal averageSimilarity = SimilarityTool.ComapareSimilarLists(totalSubjsOfsourceWS,
            //        oktmoCurrColumnList);
            //    columnMatchDictionary.Add(i, averageSimilarity);
            //    if (averageSimilarity > percentIsOk) break;
            //}

            //if (columnMatchDictionary.Count == 0) return;
            //result = columnMatchDictionary.Aggregate((l, r) => l.Value > r.Value ? l : r).Key + 1;

            //columnsDictionary.Add(result, columnCode);
            //checkedColumnsList.Add(result);
            //JustColumn firstOrDefault = sourceColumns.FirstOrDefault(x => x.Index == result - 1);
            //if (firstOrDefault != null)
            //    firstOrDefault.Code = columnCode;
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
//                    JustColumn firstOrDefault = sourceColumns.FirstOrDefault(x => x.Index == i);
//                    if (firstOrDefault != null)
//                        firstOrDefault.CodeName = columnCode;
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
            //var landCategorioesVersionsList = new List<string>
            //{
            //    "земли сельскохозяйственного назначения",
            //    "земли населенных пунктов",
            //    "земли промышленности и иного назначения",
            //    "земли особо охраняемых территорий и объектов",
            //    "земли лесного фонда",
            //    "земли водного фонда",
            //    "земли запаса"
            //};
            ////landCategorioesVersionsList.Add("не указано");

            //var maskList = new List<string>{"категор", "земл"};
            //if (GetColumnNumberByColumnName("LAND_CATEGORY", maskList)) return;

            //Dictionary<int, decimal> columnMatchDictionary = new Dictionary<int, decimal>();

            //const decimal percentIsOk = (decimal) 0.55;

            ////Берём столбец
            //for (int i = 0; i < wsTable.Columns.Count; i++)
            //{
            //    if (checkedColumnsList.Contains(i + 1)) continue;
            //    if (wsTable.Columns[i].DataType != Type.GetType("System.String")) continue;


            //    //Вce объекты
            //    List<string> totalSubjsOfsourceWS = (from row in wsTable.AsEnumerable()
            //        where !String.IsNullOrEmpty(row.Field<string>(i))
            //        select row.Field<String>(i)).ToList();
            //    if (totalSubjsOfsourceWS.Any(s => s.Contains("http") || s.Length > 100)) continue;

            //    //Получаем процент схожести между 
            //    decimal averageSimilarity = SimilarityTool.ComapareSimilarLists(totalSubjsOfsourceWS,
            //        landCategorioesVersionsList);
            //    columnMatchDictionary.Add(i, averageSimilarity);
            //    if (averageSimilarity > percentIsOk) break;
            //}
            //if (columnMatchDictionary.Count == 0) return;
            ////Берёт наибольшее значение
            //int result = columnMatchDictionary.Aggregate((l, r) => l.Value > r.Value ? l : r).Key + 1;

            //columnsDictionary.Add(result, "LAND_CATEGORY");
            //checkedColumnsList.Add(result);
            //JustColumn firstOrDefault = sourceColumns.FirstOrDefault(x => x.Index == result - 1);
            //if (firstOrDefault != null)
            //    firstOrDefault.Code = "LAND_CATEGORY";
        }

        private void GetPERMITTED_USE()
        {
            List<string> maskList = new List<string>(new[] { "ВИД_РАЗРЕШЕННОГО_ИСПОЛЬЗОВАНИЯ", "вид р", "разрешен", "использ" });

            string columnCode = "PERMITTED_USE";
            if (GetColumnNumberByColumnName(columnCode, maskList)) return;

            //List<string> permissionUseVersionsList = new List<string>
            //{
            //    "для размещения объектов сельскохозяйственного назначения и сельскохозяйственных угодий",
            //    "для сельскохозяйственного производства",
            //    "для ведения крестьянского (фермерского) хозяйства",
            //    "для ведения личного подсобного хозяйства",
            //    "для ведения гражданами садоводства и огородничества",
            //    "для ведения гражданами животноводства",
            //    "для дачного строительства",
            //    "для сенокошения и выпаса скота гражданами",
            //    "для размещения объектов охотничьего хозяйства",
            //    "для размещения объектов рыбного хозяйства",
            //    "для иных видов сельскохозяйственного использования",
            //    "для размещения объектов, характерных для населенных пунктов",
            //    "для объектов жилой застройки",
            //    "для индивидуальной жилой застройки",
            //    "для многоквартирной застройки",
            //    "для размещения объектов дошкольного, начального, общего и среднего (полного) общего образования",
            //    "для размещения иных объектов, допустимых в жилых зонах и не перечисленных в классификаторе",
            //    "для объектов общественно-делового значения",
            //    "для размещения объектов социального и коммунально-бытового назначения",
            //    "для размещения объектов здравоохранения",
            //    "для размещения объектов культуры",
            //    "для размещения объектов торговли",
            //    "для размещения объектов общественного питания",
            //    "для размещения объектов предпринимательской деятельности",
            //    "для размещения объектов среднего профессионального и высшего профессионального образования",
            //    "для размещения административных зданий",
            //    "для размещения научно-исследовательских учреждений",
            //    "для размещения культовых зданий",
            //    "для стоянок автомобильного транспорта",
            //    "для размещения объектов делового назначения, в том числе офисных центров",
            //    "для размещения объектов финансового назначения",
            //    "для размещения гостиниц",
            //    "для размещения подземных или многоэтажных гаражей",
            //    "для размещения индивидуальных гаражей",
            //    "для размещения иных объектов общественно-делового значения, обеспечивающих жизнь граждан",
            //    "для общего пользования (уличная сеть)",
            //    "для размещения объектов специального назначения",
            //    "для размещения коммунальных, складских объектов",
            //    "для размещения объектов жилищно-коммунального хозяйства",
            //    "для иных видов использования, характерных для населенных пунктов",
            //    "для размещения объектов промышленности, энергетики, транспорта, связи, радиовещания, телевидения, информатики",
            //    "для размещения промышленных объектов",
            //    "для размещения производственных и административных зданий, строений, сооружений и обслуживающих их объектов",
            //    "для размещения объектов энергетики",
            //    "для размещения объектов транспорта",
            //    "для размещения объектов связи, радиовещания, телевидения, информатики",
            //    "для размещения объектов, предназначенных для обеспечения космической деятельности",
            //    "для размещения объектов, предназначенных для обеспечения обороны и безопасности"
            //};

            ////permissionUseVersionsList.Add("не указано");

            //Dictionary<int, decimal> columnMatchDictionary = new Dictionary<int, decimal>();
            //const decimal percentIsOk = (decimal) 0.55;

            


            ////Берём столбец
            //for (int i = 0; i < wsTable.Columns.Count; i++)
            //{
            //    if (checkedColumnsList.Contains(i + 1)) continue;
            //    if (wsTable.Columns[i].DataType != Type.GetType("System.String")) continue;

            //    //Вce объекты
            //    List<string> totalSubjsOfsourceWS = (from row in wsTable.AsEnumerable()
            //        where !String.IsNullOrEmpty(row.Field<string>(i))
            //        select row.Field<String>(i)).ToList();
            //    if (totalSubjsOfsourceWS.Any(s => s.Contains("http") || s.Length > 100)) continue;
            //    //Убираем дубли
            //    List<string> uniqSubjsOfsourceWs = totalSubjsOfsourceWS.Distinct().ToList();

            //    //Вес объекта
            //    Dictionary<string, decimal> weigthOfObjects =
            //        totalSubjsOfsourceWS.GroupBy(x => x)
            //            .ToDictionary(x => x.Key,
            //                x => Decimal.Round((decimal) x.Count()/(decimal) totalSubjsOfsourceWS.Count, 4));
            //    //Console.WriteLine(weigthOfObjects.Values.Sum(x => x));
            //    //Словарь где записывается результат  - процент максимального совпадения с одним из пунктов LandCategoriesList
            //    //Мы получим ответ, на сколько каждая строка из уникальных в выгрузке соответствует максимально похожему критерию Категория земли
            //    Dictionary<string, decimal> similaritiesDictionary = new Dictionary<string, decimal>();
            //    foreach (string s1 in uniqSubjsOfsourceWs)
            //    {
            //        List<decimal> similaritiesCellWithLandListList =
            //            permissionUseVersionsList.Select(s2 => (decimal) SimilarityTool.CompareStrings(s1, s2)).ToList();
            //        similaritiesDictionary.Add(s1, similaritiesCellWithLandListList.Max());
            //    }

            //    //decimal maxSimilarity = similaritiesDictionary.Values.Max();
            //    decimal averageSimilarity = weigthOfObjects.Select(x => similaritiesDictionary[x.Key]*x.Value).Sum();
            //    columnMatchDictionary.Add(i, averageSimilarity);
            //    if (averageSimilarity > percentIsOk) break;
            //}
            //if (columnMatchDictionary.Count == 0) return;
            //int result = columnMatchDictionary.Aggregate((l, r) => l.Value > r.Value ? l : r).Key + 1;

            //columnsDictionary.Add(result, columnCode);
            //checkedColumnsList.Add(result);
            //JustColumn firstOrDefault = sourceColumns.FirstOrDefault(x => x.Index == result - 1);
            //if (firstOrDefault != null)
            //    firstOrDefault.Code = columnCode;
        }

        private void GetSOURCE_LINK()
        {

            List<string> maskList = new List<string>(new[] { "ССЫЛКА_НА_ИСТОЧНИК_ИНФОРМАЦИИ","ССЫЛКА" });

            string columnCode = String.Empty;
//            if (templateWorkbook is LandPropertyTemplateWorkbook) 
//                columnCode = "URL_INFO";
//            else 
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

//                JustColumn firstOrDefault = sourceColumns.FirstOrDefault(x => x.Index == i);
//                if (firstOrDefault != null)
//                    firstOrDefault.CodeName = columnCode;
                
                columnsDictionary.Add(result, columnCode);
                checkedColumnsList.Add(i + 1);
                return;
            }
        }

        private void GetDIST_REG_CENTER()
        {
            List<string> mask = new List<string> {"УДАЛЕННОСТЬ", "центр"};
            GetColumnNumberByColumnName("DIST_REG_CENTER", mask);

            //foreach (DataColumn column in wsTable.Columns)
            //{
            //    if (checkedColumnsList.Contains(wsTable.Columns.IndexOf(column) + 1)) continue;
            //    if (column.DataType != Type.GetType("System.String")) continue;
            //    if (column.ColumnName.Contains(mask))
            //    {
            //        columnsDictionary.Add(wsTable.Columns.IndexOf(column) + 1, "DIST_REG_CENTER");
            //        checkedColumnsList.Add(wsTable.Columns.IndexOf(column) + 1);
            //        JustColumn firstOrDefault =
            //            sourceColumns.FirstOrDefault(x => x.Index == wsTable.Columns.IndexOf(column));
            //        if (firstOrDefault != null)
            //            firstOrDefault.Code = "DIST_REG_CENTER";
            //        break;
            //    }
            //}
        }

        private void GetSYSTEM_GAS()
        {
            List<string> maskList = new List<string> {"ГАЗОСНАБЖЕНИЕ","газоснаб", "газос", "газ", "коммуникац", "коммуник", "комм"};
            if (GetColumnNumberByColumnName("SYSTEM_GAS", maskList)) return;
            //foreach (DataColumn column in wsTable.Columns)
            //{
            //    if (checkedColumnsList.Contains(wsTable.Columns.IndexOf(column) + 1)) continue;
            //    if (column.DataType != Type.GetType("System.String")) continue;
            //    if (column.ColumnName.IndexOf(mask, StringComparison.OrdinalIgnoreCase) >= 0)
            //    {
            //        columnsDictionary.Add(wsTable.Columns.IndexOf(column) + 1, "SYSTEM_GAS");
            //        checkedColumnsList.Add(wsTable.Columns.IndexOf(column) + 1);
            //        JustColumn firstOrDefault =
            //            sourceColumns.FirstOrDefault(x => x.Index == wsTable.Columns.IndexOf(column));
            //        if (firstOrDefault != null)
            //            firstOrDefault.Code = "SYSTEM_GAS";
            //        break;
            //    }
            //}
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
            //foreach (DataColumn column in wsTable.Columns)
            //{
            //    if (checkedColumnsList.Contains(wsTable.Columns.IndexOf(column) + 1)) continue;
            //    if (column.DataType != Type.GetType("System.String")) continue;
            //    if (column.ColumnName.IndexOf(mask, StringComparison.OrdinalIgnoreCase) >= 0)
            //    {
            //        columnsDictionary.Add(wsTable.Columns.IndexOf(column) + 1, "SYSTEM_SEWERAGE");
            //        checkedColumnsList.Add(wsTable.Columns.IndexOf(column) + 1);
            //        JustColumn firstOrDefault =
            //            sourceColumns.FirstOrDefault(x => x.Index == wsTable.Columns.IndexOf(column));
            //        if (firstOrDefault != null)
            //            firstOrDefault.Code = "SYSTEM_SEWERAGE";
            //        break;
            //    }
            //}

        }

        private void GetSYSTEM_ELECTRICITY()
        {

            List<string> maskList = new List<string> {"ЭЛЕКТРОСНАБЖЕНИЕ", "электроснаб", "электрос", "электро","эле" };
            if (GetColumnNumberByColumnName("SYSTEM_ELECTRICITY", maskList)) return;
            //string mask = "элект";
            //foreach (DataColumn column in wsTable.Columns)
            //{
            //    if (checkedColumnsList.Contains(wsTable.Columns.IndexOf(column) + 1)) continue;
            //    if (column.DataType != Type.GetType("System.String")) continue;
            //    if (column.ColumnName.IndexOf(mask, StringComparison.OrdinalIgnoreCase) >= 0)
            //    {
            //        columnsDictionary.Add(wsTable.Columns.IndexOf(column) + 1, "SYSTEM_ELECTRICITY");
            //        checkedColumnsList.Add(wsTable.Columns.IndexOf(column) + 1);
            //        JustColumn firstOrDefault =
            //            sourceColumns.FirstOrDefault(x => x.Index == wsTable.Columns.IndexOf(column));
            //        if (firstOrDefault != null)
            //            firstOrDefault.Code = "SYSTEM_ELECTRICITY";
            //        break;
            //    }
            //}
        }

        private void GetHEAT_SUPPLY()
        {

            List<string> maskList = new List<string> {"ТЕПЛОСНАБЖЕНИЕ", "теплоснаб", "тепл", "обогр", "отопл" };
            if (GetColumnNumberByColumnName("HEAT_SUPPLY", maskList)) return;
            //string mask = "теплоснаб";
            //foreach (DataColumn column in wsTable.Columns)
            //{
            //    if (checkedColumnsList.Contains(wsTable.Columns.IndexOf(column) + 1)) continue;
            //    if (column.DataType != Type.GetType("System.String")) continue;
            //    if (column.ColumnName.IndexOf(mask, StringComparison.OrdinalIgnoreCase) >= 0)
            //    {
            //        columnsDictionary.Add(wsTable.Columns.IndexOf(column) + 1, "HEAT_SUPPLY");
            //        checkedColumnsList.Add(wsTable.Columns.IndexOf(column) + 1);
            //        JustColumn firstOrDefault =
            //            sourceColumns.FirstOrDefault(x => x.Index == wsTable.Columns.IndexOf(column));
            //        if (firstOrDefault != null)
            //            firstOrDefault.Code = "HEAT_SUPPLY";
            //        break;
            //    }
            //}

        }

        private void GetRELIEF()
        {
            List<string> maskList = new List<string> { "рельеф"};
            if (GetColumnNumberByColumnName("RELIEF", maskList)) return;
            //string mask = "рельеф";
            //foreach (DataColumn column in wsTable.Columns)
            //{
            //    if (checkedColumnsList.Contains(wsTable.Columns.IndexOf(column) + 1)) continue;
            //    if (column.DataType != Type.GetType("System.String")) continue;
            //    if (column.ColumnName.IndexOf(mask, StringComparison.OrdinalIgnoreCase) >= 0)
            //    {
            //        columnsDictionary.Add(wsTable.Columns.IndexOf(column) + 1, "RELIEF");
            //        checkedColumnsList.Add(wsTable.Columns.IndexOf(column) + 1);
            //        JustColumn firstOrDefault =
            //            sourceColumns.FirstOrDefault(x => x.Index == wsTable.Columns.IndexOf(column));
            //        if (firstOrDefault != null)
            //            firstOrDefault.Code = "RELIEF";
            //        break;
            //    }
            //}
        }

        private void GetVEGETATION()
        {

            List<string> maskList = new List<string> { "растен" };
            if (GetColumnNumberByColumnName("VEGETATION", maskList)) return;
            //string mask = "растен";
            //foreach (DataColumn column in wsTable.Columns)
            //{
            //    if (checkedColumnsList.Contains(wsTable.Columns.IndexOf(column) + 1)) continue;
            //    if (column.DataType != Type.GetType("System.String")) continue;
            //    if (column.ColumnName.IndexOf(mask, StringComparison.OrdinalIgnoreCase) >= 0)
            //    {
            //        columnsDictionary.Add(wsTable.Columns.IndexOf(column) + 1, "VEGETATION");
            //        checkedColumnsList.Add(wsTable.Columns.IndexOf(column) + 1);
            //        JustColumn firstOrDefault =
            //            sourceColumns.FirstOrDefault(x => x.Index == wsTable.Columns.IndexOf(column));
            //        if (firstOrDefault != null)
            //            firstOrDefault.Code = "VEGETATION";
            //        break;
            //    }
            //}
        }

        private void GetPRICE()
        {
            List<string> maskList = new List<string> { "СТОИМОСТЬ", "стоим", "цена", "продаж" };
            if (templateWorkbook is LandPropertyTemplateWorkbook)
                GetColumnNumberByColumnName("PRICE", maskList);
            else
                GetColumnNumberByColumnName("SALE_PRICE", maskList);
                
            //List<string> masklList = new List<string> {"стоим", "цен"};
            //foreach (DataColumn column in from DataColumn column in wsTable.Columns
            //    where !checkedColumnsList.Contains(wsTable.Columns.IndexOf(column) + 1)
            //    //where column.DataType == Type.GetType("System.Double")
            //    where (column.ColumnName.IndexOf(masklList[0], StringComparison.OrdinalIgnoreCase) >= 0) ||
            //          (column.ColumnName.IndexOf(masklList[1], StringComparison.OrdinalIgnoreCase) >= 0)
            //    select column)
            //{
            //    columnsDictionary.Add(wsTable.Columns.IndexOf(column) + 1, "PRICE");
            //    checkedColumnsList.Add(wsTable.Columns.IndexOf(column) + 1);
            //    JustColumn firstOrDefault = sourceColumns.FirstOrDefault(x => x.Index == wsTable.Columns.IndexOf(column));
            //    if (firstOrDefault != null)
            //        firstOrDefault.Code = "PRICE";
            //    break;
            //}
        }

        private void GetDATE_RESEARCH()
        {
            if (columnsDictionary.ContainsValue("DATE_RESEARCH")) return;
            if (GetColumnNumberByColumnName("DATE_RESEARCH", new List<string> { "ДАТА_РАЗМЕЩЕНИЯ_ИНФОРМАЦИИ", "ДАТА_РАЗМЕЩЕНИЯ", "дата" })) return;

            //foreach (DataColumn column in from DataColumn column in wsTable.Columns
            //    where !checkedColumnsList.Contains(wsTable.Columns.IndexOf(column) + 1)
            //    where column.DataType == Type.GetType("System.DateTime")
            //    select column)
            //{
            //    columnsDictionary.Add(wsTable.Columns.IndexOf(column) + 1, "DATE_RESEARCH");
            //    checkedColumnsList.Add(wsTable.Columns.IndexOf(column) + 1);
            //    JustColumn firstOrDefault = sourceColumns.FirstOrDefault(x => x.Index == wsTable.Columns.IndexOf(column));
            //    if (firstOrDefault != null)
            //        firstOrDefault.Code = "DATE_RESEARCH";
            //    break;
            //}
        }

        private void GetAREA_LOT()
        {

            List<string> maskList = new List<string> {"ПЛОЩАДЬ УЧАСТКА", "ПЛОЩАДЬ_УЧАСТКА","ПЛОЩАДЬ", "площад", "площ" };
            if (GetColumnNumberByColumnName("AREA_LOT", maskList)) return;

            //const string mask = "площад";
            //foreach (DataColumn column in from DataColumn column in wsTable.Columns
            //    where !checkedColumnsList.Contains(wsTable.Columns.IndexOf(column) + 1)
            //    //where column.DataType == Type.GetType("System.Double") 
            //    where column.ColumnName.IndexOf(mask, StringComparison.OrdinalIgnoreCase) >= 0
            //    select column)
            //{

            //    columnsDictionary.Add(wsTable.Columns.IndexOf(column) + 1, "AREA_LOT");
            //    checkedColumnsList.Add(wsTable.Columns.IndexOf(column) + 1);
            //    JustColumn firstOrDefault = sourceColumns.FirstOrDefault(x => x.Index == wsTable.Columns.IndexOf(column));
            //    if (firstOrDefault != null)
            //        firstOrDefault.Code = "AREA_LOT";
            //}
        }

        private void GetCONTACTS()
        {
            GetColumnNumberByColumnName("CONTACTS", new List<string> { "ТЕЛЕФОН_ПРОДАВЦА" });
            GetColumnNumberByColumnName("CONTACTS", new List<string> { "КОНТАКТЫ" });
            GetColumnNumberByColumnName("CONTACTS", new List<string> { "ТЕЛЕФОН"});

//            Regex phoneRegex = new Regex(@"\d+[-]?\d+[-]?\d+[-]?\d+", RegexOptions.IgnoreCase);
//            Regex emailRegex = new Regex(@".+@.+\..+", RegexOptions.IgnoreCase);
//            const decimal isOk = 0.4M;
//            foreach (DataColumn column in from DataColumn column in wsTable.Columns
//                where !checkedColumnsList.Contains(wsTable.Columns.IndexOf(column) + 1)
//                where column.DataType == Type.GetType("System.String")
//                where wsTable.AsEnumerable().All(s => String.IsNullOrEmpty(s.Field<string>(column)) ||
//                                                      (!s.Field<string>(column).Contains("http") ||
//                                                       s.Field<string>(column).Length < 100 ||
//                                                       !(new Regex(@"\d\d\.\d\d\.(\d)+").IsMatch(s.Field<string>(column)))))
//                select column)
//            {
//                
//                if (column.ColumnName.IndexOf("владел", StringComparison.OrdinalIgnoreCase) >= 0 ||
//                    column.ColumnName.IndexOf("ФИО", StringComparison.OrdinalIgnoreCase) >= 0)
//                {
//                    columnsDictionary.Add(wsTable.Columns.IndexOf(column) + 1, "CONTACTS");
//                    checkedColumnsList.Add(wsTable.Columns.IndexOf(column) + 1);
//                }
//                else
//                {
//                    int totalObjects = wsTable.AsEnumerable().Count();
//                    int countPhones =
//                        wsTable.AsEnumerable()
//                            .Count(
//                                x =>
//                                    !String.IsNullOrEmpty(x.Field<string>(column)) &&
//                                    phoneRegex.IsMatch(x.Field<string>(column)));
//                    int countEmails =
//                        wsTable.AsEnumerable()
//                            .Count(
//                                x =>
//                                    !String.IsNullOrEmpty(x.Field<string>(column)) &&
//                                    emailRegex.IsMatch(x.Field<string>(column)));
//                    if (countEmails == 0 && countPhones == 0) continue;
//                    int rs = countPhones > countEmails ? countPhones : countEmails;
//
//                    if ((((decimal) rs)/((decimal) totalObjects)) <= isOk) continue;
//
//                    columnsDictionary.Add(wsTable.Columns.IndexOf(column) + 1, "CONTACTS");
//                    checkedColumnsList.Add(wsTable.Columns.IndexOf(column) + 1);
//                    JustColumn firstOrDefault =
//                        sourceColumns.FirstOrDefault(x => x.Index == wsTable.Columns.IndexOf(column));
//                    if (firstOrDefault != null)
//                        firstOrDefault.CodeName = "CONTACTS";
//                    if (columnsDictionary.AsEnumerable().Count(x => x.Value == "CONTACTS") > 2) return;
//                }
//            }
        }

        private void GetOperationType()
        {
            List<string> masksList = new List<string> {"ВИД_СДЕЛКИ","ВИД СДЕЛКИ","продажа", "аренда"};

            if (GetColumnNumberByColumnName("OPERATION", masksList)) return;

            //Dictionary<int, decimal> columnMatchDictionary = new Dictionary<int, decimal>();

            //const decimal percentIsOk = (decimal) 0.6;

            ////Берём столбец
            //for (int i = 0; i < wsTable.Columns.Count; i++)
            //{
            //    if (columnsDictionary.ContainsKey(i + 1)) continue;
            //    if (checkedColumnsList.Contains(i + 1)) continue;
            //    if (wsTable.Columns[i].DataType != Type.GetType("System.String")) continue;


            //    //Вce объекты
            //    List<string> totalSubjsOfsourceWS = (from row in wsTable.AsEnumerable()
            //        where !String.IsNullOrEmpty(row.Field<string>(i))
            //        select row.Field<String>(i)).ToList();
            //    if (totalSubjsOfsourceWS.Any(s => s.Contains("http") || s.Length > 50)) continue;

            //    //Получаем процент схожести между 
            //    decimal averageSimilarity = SimilarityTool.ComapareSimilarLists(totalSubjsOfsourceWS,
            //        masksList);
            //    columnMatchDictionary.Add(i, averageSimilarity);
            //    if (averageSimilarity > percentIsOk) break;
            //}

            //if (columnMatchDictionary.Count == 0) return;
            //int result = columnMatchDictionary.Aggregate((l, r) => l.Value > r.Value ? l : r).Key + 1;

            //columnsDictionary.Add(result, "OPERATION");
            //checkedColumnsList.Add(result);
            //JustColumn firstOrDefault = sourceColumns.FirstOrDefault(x => x.Index == result);
            //if (firstOrDefault != null)
            //    firstOrDefault.Code = "OPERATION";

        }       

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
        /// формирует список столбцов с названиями, порядковым номером и образцами содержания
        /// </summary>
        private void CreateColumnsList()
        {
//            int i = 1;
//            foreach (DataColumn column in wsTable.Columns)
//            {
//                sourceColumns.Add(new JustColumn(column.ColumnName, i)
//                {
//                    Examples = wsTable.Rows.Cast<DataRow>().Where(x => x[column] != null).Where(x => x[column].ToString() != "")
//                                    .Select(x => x[column].ToString()).ToList()
//                });
//                i++;
//            }
        }

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


            //Тоже самое но в списко колонок
//            JustColumn firstOrDefault = sourceColumns.First(x => x.Index == cl.Ordinal + 1);
//            firstOrDefault.CodeName = columnCode;

            //Результат работы
            return true;
        }

        //public Dictionary<int, int> ResultDictionary { get; set; }

        //public void SetResultDictionary(Dictionary<string, string> dictionary)
        //{
        //    Dictionary<int, int> tmpDict = new Dictionary<int, int>();
        //    foreach (KeyValuePair<string, string> keyValuePair in dictionary)
        //    {
        //        JustColumn srColumn = sourceColumns.First(x => x.Name == keyValuePair.Key);
        //        JustColumn targetColumn =LandPropertyTemplateWorkbook.TemplateColumns.First(x => x.Code == keyValuePair.Value);

        //        tmpDict.Add(srColumn.Index+1,targetColumn.Index);
        //    }
        //    //<источник/база>
        //    ResultDictionary = tmpDict;
        //}
    }
}
// ReSharper restore SuggestUseVarKeywordEvident