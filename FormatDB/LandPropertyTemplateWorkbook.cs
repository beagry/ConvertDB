using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;

namespace Formater
{
    public static class LandPropertyTemplateWorkbook
    {
        private static readonly List<JustColumn> columns = new List<JustColumn>
            #region Columns Initialize
        {
            new JustColumn("OBJECTID","Уникальный итдентификатор объекта",1),
            new JustColumn("SUBJECT", "Субъект Российской Федерации",2),
            new JustColumn("REGION", "Муниципальное образование (район)",3),
            new JustColumn("SETTLEMENT", "Поселение",4),
            new JustColumn("NEAR_CITY", "Ближайший населенный пункт",5),
            new JustColumn("TERRITORY_TYPE", "Тип ближайшего населенного пункта",6),
            new JustColumn("IN_CITY", "Объект расположен в границах населенного пункта",7),
            new JustColumn("VGT", "Городской район",8),
            new JustColumn("STREET", "Наименование адресного объекта",9),
            new JustColumn("STREET_TYPE", "Тип адресного объекта",10),
            new JustColumn("HOUSE_NUM", "Дом",11),
            new JustColumn("LETTER", "Литера",12),
            new JustColumn("BUILDING", "Корпус",13),
            new JustColumn("STRUCTURE", "Строение",14),
            new JustColumn("ESTATE", "Владение",15),
            new JustColumn("LONGITUDE", "Долгота",16),
            new JustColumn("LATITUDE", "Широта",17),
            new JustColumn("HIGHWAY", "Трасса",18),
            new JustColumn("DIST_REG_CENTER", "Расстояние до регионального центра",19),
            new JustColumn("DIST_NEAR_CITY", "Расстояние до ближайшего населенного пункта",20),
            new JustColumn("CADASTRE_NUM", "Кадастровый номер земельного участка",21),
            new JustColumn("OFFER_DEAL", "Предложение (сделка)",22),
            new JustColumn("OPERATION", "Операция",23),
            new JustColumn("LAW_NOW", "Права на участок",24),
            new JustColumn("SALE_TYPE", "Способ реализации",25),
            new JustColumn("RENTAL_PERIOD", "Срок аренды",26),
            new JustColumn("PRICE", "Цена предложения (сделки)",7), /////////////////
            new JustColumn("PRICE_FOR_UNIT", "Цена за сотку",8), /////////////
            new JustColumn("RENT_RATE", "Арендная плата",28),
            new JustColumn("AREA_LOT", "Площадь",9), ///////////////
            new JustColumn("LAND_CATEGORY", "Категория земель",30),
            new JustColumn("PERMITTED_USE", "Вид разрешенного использования",31),
            new JustColumn("PERMITTED_USE_TEXT", "Вид разрешенного использования текст",32),
            new JustColumn("SYSTEM_GAS", "Газоснабжение",33),
            new JustColumn("SYSTEM_WATER", "Водоснабжение",34),
            new JustColumn("SYSTEM_SEWERAGE", "Канализация",35),
            new JustColumn("SYSTEM_ELECTRICITY", "Электроснабжение",36),
            new JustColumn("HEAT_SUPPLY", "Теплоснабжение",37),
            new JustColumn("OBJECT", "Наличие объектов на участке",38),
            new JustColumn("SURFACE", "Покрытие площадки",39),
            new JustColumn("ROAD", "Дорога",40),
            new JustColumn("RELIEF", "Рельеф",41),
            new JustColumn("VEGETATION", "Растительный покров",42),
            new JustColumn("DESCRIPTION", "Описание",43),
            new JustColumn("SOURCE_DESC", "Источник информации",44),
            new JustColumn("URL_SALE", "Ссылка на источник информации",45),
            new JustColumn("SELLER", "Наименование продавца",46),
            new JustColumn("OKOPF", "Организационно-правовая форма",47),
            new JustColumn("URL_INFO", "Адрес сайта в сети интернет",48),
            new JustColumn("CONTACTS", "Контакты",49),
            new JustColumn("DATE_RESEARCH", "Дата размещения информации",27), ///////
            new JustColumn("DATE_IN_BASE", "Дата отчета по этапу",28), ///////
            new JustColumn("ACTUAL", "Актуальность",52),
            new JustColumn("DATE_IS_RINGING", "Дата прозвона",53),
            new JustColumn("RESULT", "Результат прозвона",54),
            new JustColumn("COMMENT", "Комментарий",55),
            new JustColumn("ADDITIONAL", "Уточненные (дополненные) характеристики",56),
            new JustColumn("ASSOCIATIONS", "Товарищества и корпоративы",57),
            new JustColumn("DATE_PARSING", "Дата парсинга", 29), //////
        };

        #endregion

        public static IEnumerable<JustColumn> TemplateColumns
        {
            get { return columns; }
        }

        public static int GetColumnByCode(string name)
        {
            int column = 0;
            JustColumn firstOrDefault = columns.FirstOrDefault(x => x.Code == name);
            if (firstOrDefault != null)
                column = firstOrDefault.Index;
            return column;
        }

        public static Dictionary<string, int> GroupWorkBooksByHead(IEnumerable<string> workbooksPaths)
        {
            var xlApplication = GetExcelApplication();
            var resultDictionary = new Dictionary<string, int>();


            var wsTypes = new List<WSType>();
            var n = 1;
            foreach (var s in workbooksPaths)
            {
                Process.Start(s);
                var workbook = xlApplication.Workbooks.Cast<Excel.Workbook>()
                    .First(x => x.Name == Path.GetFileName(s));

#if DEBUG
                Debug.Assert(workbook != null);
#endif

                Excel.Worksheet worksheet = workbook.Worksheets[1];
                var head = new List<string>();

                var lastUsedColumn = worksheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Column;
                var headRow = worksheet.Range[worksheet.Cells[1, 1], worksheet.Cells[1, lastUsedColumn]];

                foreach (Excel.Range cell in headRow)
                    if (!String.IsNullOrEmpty(cell.Value2))
                        head.Add(cell.Value2.ToString());

                if (wsTypes.Any(x => x.Heads.SequenceEqual(head)))
                {
                    resultDictionary.Add(s, wsTypes.First(x => x.Heads.SequenceEqual(head)).GroupNumber);
                }
                else
                {
                    wsTypes.Add(new WSType { Heads = head, GroupNumber = n });
                    resultDictionary.Add(s, n);
                    n++;
                }
                workbook.Close(false);
            }
            return resultDictionary;
        }

        public static Excel.Application GetExcelApplication()
        {
#if (!DEBUG)
            return new Excel.Application { Visible = false, ScreenUpdating =  false};
#endif
            Excel.Application xlApplication = null;
            try
            {
                xlApplication = (Excel.Application)Marshal.GetActiveObject("Excel.Application");
            }
            catch (COMException exception)
            {
                if (xlApplication == null)
                {
                    xlApplication = new Excel.Application() { Visible = false };
                }
                else
                {
                    throw;
                }
            }
            return xlApplication;
        }

        [Obsolete("Метод не готов!",true)]
        public static Excel.Workbook GetTemplateWorkbook()
        {
            throw new Exception("Метод не готов!");
            //Excel.Workbook workbook = (new Excel.Application()).Workbooks.Add();
            //return workbook;
        }
        
    }

    public class JustColumn
    {
        public JustColumn(string name, string description, int index)
        {
            Index = index;
            Name = description;
            Code = name;
        }

        public JustColumn(string description, int index)
        {
            Index = index;
            Name = description;
        }
        public int Index { get; set; }

        public string Name { get; private set; }

        public string Code { get; set; }

        public List<string> Examples { get; set; }
    }
    public class WSType
    {
        public List<string> Heads { get; set; }
        public int GroupNumber { get; set; }

    }
}
