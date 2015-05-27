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
        private const int TakeFirstItemsQuantity = 300;
        private readonly List<int> checkedColumnsList;

        /// <summary>
        ///     Key = номер столбца, который будет скопирован, Value = Название колонки Куда будет скопирован столбец
        /// </summary>
        private readonly Dictionary<int, string> columnsDictionary = new Dictionary<int, string>();

        private readonly Dictionary<int, string> head;
        private readonly TemplateWorkbook templateWorkbook;
        private readonly Template_workbooks.EFModels.TemplateWorkbook wb;
        private readonly DataTable wsTable;


        /// <summary>
        ///     Старый констуктор для UpdateWB проекта
        /// </summary>
        /// <param name="worksheet"></param>
        /// <param name="templateWorkbook"></param>
        public SourceWs(Worksheet worksheet, TemplateWorkbook templateWorkbook) : this(templateWorkbook)
        {
            var sourceWorksheet = worksheet;

            wsTable = FillDataTable.GetDataTable(((Workbook) sourceWorksheet.Parent).FullName,
                sourceWorksheet.Name, TakeFirstItemsQuantity);
            head = worksheet.ReadHead();
        }

        private SourceWs(TemplateWorkbook workbook) : this()
        {
            templateWorkbook = workbook;
        }

        /// <summary>
        ///     Самый продуктивный конструктор
        /// </summary>
        /// <param name="table"></param>
        /// <param name="wbType"></param>
        public SourceWs(DataTable table, XlTemplateWorkbookType wbType)
            : this()
        {
            wsTable = table;
            head = wsTable.Columns.Cast<DataColumn>()
                .ToDictionary(k => wsTable.Columns.IndexOf(k) + 1, v => v.ColumnName);

            var db = TemplateWbsRepository.Context;
            wb = db.TemplateWorkbooks.First(w => w.WorkbookType == wbType);
        }

        private SourceWs()
        {
            checkedColumnsList = new List<int>();
        }

        public Dictionary<string, List<string>> ResultDictionary
        {
            get
            {
                return columnsDictionary
                    .Select(kp => new {ColumnCopy = head.First(hk => hk.Key == kp.Key).Value, ColumnPaste = kp.Value})
                    .GroupBy(obj => obj.ColumnPaste, o => o.ColumnCopy)
                    .ToDictionary(k => k.Key, v => v.ToList());
            }
        }

        public void CheckColumns()
        {
            //
            //Общие колонки
            //

            TryToFindTemplateColumnsFromDbData();
            return;

            GetDESCRIPTION();
            GetSUBJECT();
            GetDateParsing();
            GetSOURCE_DESC();

            //Уникальные поля Зем участков
            if (templateWorkbook is LandPropertyTemplateWorkbook)
            {
                GetRights();
                GetBuildings();
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

        /// <summary>
        ///     Метод находит колонку с полным или частичным совпадением в имени по пеерданному списку критериев поиска
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

            if (cl == null) return false;

            checkedColumnsList.Add(cl.Ordinal + 1);
            columnsDictionary.Add(cl.Ordinal + 1, columnCode);
            return true;
        }

        #region GetColumnMethods

        private void GetSUBJECT()
        {
            var columnMatchDictionary = new Dictionary<int, decimal>();
            int result;
            const decimal percentIsOk = (decimal) 0.4;
            var maskList =
                new List<string>(new[] {"СУБЪЕКТ_РОССИЙСКОЙ_ФЕДЕРАЦИИ", "субъект", "република", "область", "край"});
            if (GetColumnNumberByColumnName("SUBJECT", maskList)) return;
            //В каждой колонке поочередно
            for (var i = 0; i < wsTable.Columns.Count; i++)
            {
                if (checkedColumnsList.Contains(i + 1)) continue;
                if (wsTable.Columns[i].DataType != Type.GetType("System.String")) continue;

                //Берём все уникальные объекты
                var uniqSubjsOfsourceWS = (from x in wsTable.AsEnumerable()
                    where !string.IsNullOrEmpty(x.Field<string>(i))
                    select x.Field<string>(i)).Distinct().ToList();

                var cup = (int) (percentIsOk*uniqSubjsOfsourceWS.Count());
                if (uniqSubjsOfsourceWS.Any(s => s.Contains("http") || s.Length > 100)) continue;
                var fitCellsQuantity =
                    uniqSubjsOfsourceWS.Count(
                        x => x.Contains(maskList[2]) || x.Contains(maskList[3]) || x.Contains(maskList[4]));

                if (fitCellsQuantity == 0) continue;
                //1.0 = 100% значений столбца
                var resultDecimal = decimal.Round((decimal) fitCellsQuantity/uniqSubjsOfsourceWS.Count(), 2);
                //Пишем результать совпадений в колонке
                columnMatchDictionary.Add(i, resultDecimal);
                if (resultDecimal >= percentIsOk) break;
            }
            if (columnMatchDictionary.Count == 0) return;
            result = columnMatchDictionary.Aggregate((l, r) => l.Value > r.Value ? l : r).Key + 1;

            columnsDictionary.Add(result, "SUBJECT");
            checkedColumnsList.Add(result);
        }


        private void GetDESCRIPTION()
        {
            var maskList = new List<string> {"ОПИСАНИЕ"};
            const string columnCode = "DESCRIPTION";
            if (GetColumnNumberByColumnName(columnCode, maskList)) return;

            if (columnsDictionary.ContainsValue(columnCode)) return;

            int[] cups = {300, 150, 100};
            foreach (var cup in cups)
            {
                for (var i = 0; i < wsTable.Columns.Count; i++)
                {
                    if (checkedColumnsList.Contains(i + 1)) continue;
                    if (wsTable.Columns[i].DataType != Type.GetType("System.String")) continue;
                    //Просто находим столец, в котором очень много букв
                    var c = (wsTable.AsEnumerable().Where(x => !string.IsNullOrEmpty(x.Field<string>(i)))).Count(
                        x => x.Field<string>(i).Length > cup);
                    if (c == 0) continue;

                    //Нашли
                    columnsDictionary.Add(i + 1, columnCode);
                    checkedColumnsList.Add(i + 1);
                    return;
                }
            }
        }

        private void GetSOURCE_LINK()
        {
            var maskList = new List<string>(new[] {"ССЫЛКА_НА_ИСТОЧНИК_ИНФОРМАЦИИ", "ССЫЛКА"});

            const string columnCode = "SOURCE_LINK";

            if (columnCode == string.Empty) return;

            if (columnsDictionary.ContainsValue(columnCode)) return;

            if (GetColumnNumberByColumnName(columnCode, maskList)) return;


            for (var i = 0; i < wsTable.Columns.Count; i++)
            {
                if (checkedColumnsList.Contains(i + 1)) continue;
                if (wsTable.Columns[i].DataType != Type.GetType("System.String")) continue;
                if (
                    !wsTable.AsEnumerable()
                        .Where(x => !string.IsNullOrEmpty(x.Field<string>(i)))
                        .Any(x => x.Field<string>(i).Contains("http"))) continue;
                var percentSimilarity =
                    wsTable.AsEnumerable()
                        .Where(x2 => x2.Field<string>(i) != null)
                        .Count(x1 => x1.Field<string>(i).Contains("http"))/
                    (decimal) (wsTable.AsEnumerable().Where
                        (x2 => !string.IsNullOrEmpty(x2.Field<string>(i)))
                        )
                        .Count();
                if (percentSimilarity < 0.5M) continue;
                var result = i + 1;

                columnsDictionary.Add(result, columnCode);
                checkedColumnsList.Add(i + 1);
                return;
            }
        }

        //BUG Ниже идут комерческие поля

        private void GetBALCONY()
        {
            GetColumnNumberByColumnName("BALCONY", new List<string> {"ЛОДЖИЯ"});
            GetColumnNumberByColumnName("BALCONY", new List<string> {"БАЛКОН"});
        }

        private void GetOBJECT_TYPE()
        {
            GetColumnNumberByColumnName("OBJECT_TYPE", new List<string> {"ТИП_ОБЬЕКТА"});
        }

        private void GetAREA_TOTAL()
        {
            GetColumnNumberByColumnName("AREA_TOTAL", new List<string> {"ПЛОЩАДЬ ОБЪЕКТА", "ПЛОЩАДЬ_ОБЪЕКТА"});
        }


        private void GetDateParsing()
        {
            var maskList = new List<string> {"ДАТА_ПАРСИНГА"};
            GetColumnNumberByColumnName("DATE_PARSING", maskList);
        }

        private void GetFLOORITY()
        {
            var maskList = new List<string> {"ЭТАЖНОСТЬ"};
            GetColumnNumberByColumnName("FLOOR_QNT_MIN", maskList);
        }

        private void GetCLASS_TYPE()
        {
            var maskList = new List<string> {"ПОТРЕБИТЕЛЬСКИЙ_КЛАСС"};
            GetColumnNumberByColumnName("CLASS_TYPE", maskList);
        }

        private void GetYEAR_BUILD()
        {
            var maskList = new List<string> {"ГОД_ПОСТРОЙКИ"};
            GetColumnNumberByColumnName("YEAR_BUILD", maskList);
        }

        private void GetFLOOR()
        {
            var maskList = new List<string> {"ЭТАЖ"};
            GetColumnNumberByColumnName("FLOOR", maskList);
        }

        private void GetOBJECT_PURPOSE()
        {
            var maskList = new List<string> {"НАЗНАЧЕНИЕ_ОБЪЕКТА"};
            GetColumnNumberByColumnName("OBJECT_PURPOSE", maskList);
        }

        private void GetBUILDING_TYPE()
        {
            var maskList = new List<string> {"ТИП_ПОСТРОЙКИ"};
            GetColumnNumberByColumnName("BUILDING_TYPE", maskList);
        }

        private void GetSEGMENT()
        {
            var maskList = new List<string> {"СЕГМЕНТ"};
            GetColumnNumberByColumnName("SEGMENT", maskList);
        }

        private void GetSECURITY()
        {
            var maskList = new List<string> {"БЕЗОПАСНОСТЬ"};
            GetColumnNumberByColumnName("SECURITY", maskList);
        }

        private void GetCONDITION()
        {
            var maskList = new List<string> {"СОСТОЯНИЕ"};
            GetColumnNumberByColumnName("CONDITION", maskList);
        }

        private void GetMATERIAL_WALL()
        {
            var maskList = new List<string> {"МАТЕРИАЛ_СТЕН"};
            GetColumnNumberByColumnName("MATERIAL_WALL", maskList);
        }

        private void GetHEIGHT_FLOOR()
        {
            var maskList = new List<string> {"ВЫСОТА_ПОТОЛКА"};
            GetColumnNumberByColumnName("HEIGHT_FLOOR", maskList);
        }

        private void GetRights()
        {
            var maskList = new List<string> {"ВИД_ПРАВА", "ВИД ПРАВА", "права", "прав"};
            GetColumnNumberByColumnName("LAW_NOW", maskList);
        }

        private void GetBuildings()
        {
            var maskList = new List<string> {"СТРОЕНИЯ"};
            GetColumnNumberByColumnName("OBJECT", maskList);
        }

        private void GetSeller()
        {
            var maskList = new List<string> {"КОМПАНИЯ"};
            GetColumnNumberByColumnName("SELLER", maskList);
        }

        private void GetSOURCE_DESC()
        {
            var maskList = new List<string> {"ИСТОЧНИК_ИНФОРМАЦИИ", "ИСТОЧНИК"};
            GetColumnNumberByColumnName("SOURCE_DESC", maskList);
        }

        #endregion
    }
}

// ReSharper restore SuggestUseVarKeywordEvident