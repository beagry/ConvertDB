using System.Collections.Generic;
using System.Data.Entity;
using System.Linq;

namespace Converter.Template_workbooks.EFModels
{
    internal class TemplateWbsInitializer : DropCreateDatabaseIfModelChanges<TemplateWbsContext>
    {
        protected override void Seed(TemplateWbsContext context)
        {
            base.Seed(context);

            InitializeLandWorkbook(context);
            InitializeComerceWorkbook(context);
        }

        private void InitializeComerceWorkbook(TemplateWbsContext context)
        {
            #region columns
            var columns = new[]
            {
                new TemplateColumn {CodeName = "ID", Name = "ПОРЯДКОВЫЙ_НОМЕР", ColumnIndex = 1},
                new TemplateColumn
                {
                    CodeName = "SUBJECT",
                    Name = "СУБЪЕКТ_РОССИЙСКОЙ_ФЕДЕРАЦИИ",
                    ColumnIndex = 2,
                    SearchCritetias = new List<SearchCritetia>(new[]
                    {
                        "СУБЪЕКТ_РОССИЙСКОЙ_ФЕДЕРАЦИИ", "субъект", "република", "область", "край","регион"
                    }.Select(s => new SearchCritetia {Text = s}).ToList())
                },
                new TemplateColumn {CodeName = "REGION", Name = "МУНИЦИПАЛЬНОЕ_ОБРАЗОВАНИЕ_(РАЙОН)", ColumnIndex = 3,
                    SearchCritetias = new List<SearchCritetia>(new[]
                    {
                        "МЕСТОПОЛОЖЕНИЕ", "район", "город", "местоп"
                    }.Select(s => new SearchCritetia {Text = s}).ToList())},
                new TemplateColumn {CodeName = "SETTLEMENT", Name = "ПОСЕЛЕНИЕ", ColumnIndex = 4,
                    SearchCritetias = new List<SearchCritetia>(new[]
                    {
                        "населенн", "насел","город"
                    }.Select(s => new SearchCritetia {Text = s}).ToList())},
                new TemplateColumn {CodeName = "CITY", Name = "НАСЕЛЕННЫЙ_ПУНКТ", ColumnIndex = 5},
                new TemplateColumn {CodeName = "CITY_TYPE", Name = "ТИП_НАСЕЛЕННОГО_ПУНКТА", ColumnIndex = 6},
                new TemplateColumn {CodeName = "VGT", Name = "ВНУТРИГОРОДСКАЯ_ТЕРРИТОРИЯ", ColumnIndex = 7,
                    SearchCritetias = new List<SearchCritetia>(new[]
                    {
                        "внутригор","территор"
                    }.Select(s => new SearchCritetia {Text = s}).ToList())},
                new TemplateColumn {CodeName = "STREET", Name = "УЛИЦА", ColumnIndex = 8,
                    SearchCritetias = new List<SearchCritetia>(new[]
                    {
                        "улиц"
                    }.Select(s => new SearchCritetia {Text = s}).ToList())},
                new TemplateColumn {CodeName = "STREET_TYPE", Name = "ТИП_УЛИЦЫ", ColumnIndex = 9},
                new TemplateColumn {CodeName = "HOUSE_NUM", Name = "ДОМ", ColumnIndex = 10},
                new TemplateColumn {CodeName = "LETTER", Name = "ЛИТЕРА", ColumnIndex = 11},
                new TemplateColumn {CodeName = "BUILDING", Name = "КОРПУС", ColumnIndex = 12},
                new TemplateColumn {CodeName = "STRUCTURE", Name = "СТРОЕНИЕ", ColumnIndex = 13},
                new TemplateColumn {CodeName = "ESTATE", Name = "ВЛАДЕНИЕ", ColumnIndex = 14},
                new TemplateColumn {CodeName = "LONGITUDE", Name = "ДОЛГОТА", ColumnIndex = 15},
                new TemplateColumn {CodeName = "LATITUDE", Name = "ШИРОТА", ColumnIndex = 16},
                new TemplateColumn
                {
                    CodeName = "DIST_REG_CENTER",
                    Name = "УДАЛЕННОСТЬ_ОТ_РЕГИОНАЛЬНОГО_ЦЕНТРА",
                    ColumnIndex = 17,
                    SearchCritetias = new List<SearchCritetia>(new[]
                    {
                        "УДАЛЕННОСТЬ", "центр"
                    }.Select(s => new SearchCritetia {Text = s}).ToList())
                },
                new TemplateColumn
                {
                    CodeName = "CADASTRE_NUM",
                    Name = "КАДАСТРОВЫЙ_НОМЕР_ЗЕМЕЛЬНОГО_УЧАСТКА",
                    ColumnIndex = 18
                },
                new TemplateColumn {CodeName = "METRO", Name = "СТАНЦИЯ_МЕТРО", ColumnIndex = 19},
                new TemplateColumn {CodeName = "METRO_DISTMIN", Name = "ДО_МЕТРО_МИНУТ", ColumnIndex = 20},
                new TemplateColumn {CodeName = "TRANSPORT", Name = "ПЕШКОМ_ТРАНСПОРТОМ", ColumnIndex = 21},
                new TemplateColumn {CodeName = "SEGMENT", Name = "СЕГМЕНТ", ColumnIndex = 22},
                new TemplateColumn {CodeName = "BUILDING_TYPE", Name = "ТИП_ПОСТРОЙКИ", ColumnIndex = 23},
                new TemplateColumn {CodeName = "CENTER_CodeName", Name = "НАИМЕНОВАНИЕ_ЦЕНТРА", ColumnIndex = 24},
                new TemplateColumn {CodeName = "OBJECT_TYPE", Name = "ТИП_ОБЪЕКТА", ColumnIndex = 25},
                new TemplateColumn {CodeName = "OBJECT_PURPOSE", Name = "НАЗНАЧЕНИЕ_ОБЪЕКТА", ColumnIndex = 26},
                new TemplateColumn {CodeName = "CLASS_TYPE", Name = "ПОТРЕБИТЕЛЬСКИЙ_КЛАСС", ColumnIndex = 27},
                new TemplateColumn {CodeName = "OPERATION", Name = "ОПЕРАЦИЯ", ColumnIndex = 28,
                    SearchCritetias = new List<SearchCritetia>(new[]
                    {
                        "вид_сделки","вид сделки","сделк"
                    }.Select(s => new SearchCritetia {Text = s}).ToList())},
                new TemplateColumn {CodeName = "SALE_PRICE", Name = "ЦЕНА _ПРОДАЖИ", ColumnIndex = 29,
                    SearchCritetias = new List<SearchCritetia>(new[]
                    {
                        "СТОИМОСТЬ", "стоим", "цена", "продаж"
                    }.Select(s => new SearchCritetia {Text = s}).ToList())},
                new TemplateColumn {CodeName = "RENT_RATE", Name = "АРЕНДНАЯ_ПЛАТА", ColumnIndex = 30},
                new TemplateColumn {CodeName = "AREA", Name = "ПЛОЩАДЬ", ColumnIndex = 31,
                    SearchCritetias = new List<SearchCritetia>(new[]
                    {
                        "ПЛОЩАДЬ", "площад", "площ"
                    }.Select(s => new SearchCritetia {Text = s}).ToList())},
                new TemplateColumn {CodeName = "PRICE_FOR_UNIT", Name = "ЦЕНА_ЗА_М2", ColumnIndex = 32},
                new TemplateColumn {CodeName = "OPERATING_COSTS", Name = "ЭКСПЛУАТАЦИОННЫЕ_РАСХОДЫ", ColumnIndex = 33},
                new TemplateColumn {CodeName = "FLOOR", Name = "ЭТАЖ", ColumnIndex = 34},
                new TemplateColumn {CodeName = "FLOOR_QNT_MIN", Name = "ЭТАЖНОСТЬ_МИНИМАЛЬНАЯ", ColumnIndex = 35},
                new TemplateColumn {CodeName = "FLOOR_QNT_MAX", Name = "ЭТАЖНОСТЬ_МАКСИМАЛЬНАЯ", ColumnIndex = 36},
                new TemplateColumn {CodeName = "YEAR_BUILD", Name = "ГОД_ПОСТРОЙКИ", ColumnIndex = 37},
                new TemplateColumn {CodeName = "MATERIAL_WALL", Name = "МАТЕРИАЛ_СТЕН", ColumnIndex = 38},
                new TemplateColumn {CodeName = "HEIGHT_FLOOR", Name = "ВЫСОТА_ПОТОЛКА", ColumnIndex = 39},
                new TemplateColumn {CodeName = "COLUMN_DIST", Name = "ШАГ_КОЛОНН", ColumnIndex = 40},
                new TemplateColumn {CodeName = "LAYOUT", Name = "ПЛАНИРОВКА", ColumnIndex = 41},
                new TemplateColumn {CodeName = "ROOM_QNT", Name = "КОЛИЧЕСТВО_КОМНАТ", ColumnIndex = 42},
                new TemplateColumn {CodeName = "AREA_TOTAL", Name = "ПЛОЩАДЬ_ОБЩАЯ", ColumnIndex = 43},
                new TemplateColumn
                {
                    CodeName = "AREA_LOT",
                    Name = "ПЛОЩАДЬ_ЗЕМЕЛЬНОГО_УЧАСТКА_ОБЪЕКТА",
                    ColumnIndex = 44
                },
                new TemplateColumn {CodeName = "CONDITION", Name = "СОСТОЯНИЕ", ColumnIndex = 45},
                new TemplateColumn {CodeName = "SECURITY", Name = "БЕЗОПАСНОСТЬ", ColumnIndex = 46},
                new TemplateColumn {CodeName = "FLOOR_LOAD", Name = "ДОПУСТИМАЯ НАГРУЗКА НА ПОЛ", ColumnIndex = 47},
                new TemplateColumn {CodeName = "CONDITIONING", Name = "КОНДИЦИОНИРОВАНИЕ", ColumnIndex = 48},
                new TemplateColumn {CodeName = "VENT", Name = "ВЕНТИЛЯЦИЯ", ColumnIndex = 49},
                new TemplateColumn
                {
                    Name = "ГАЗОСНАБЖЕНИЕ",
                    CodeName = "SYSTEM_GAS",
                    ColumnIndex = 50,
                    SearchCritetias = new List<SearchCritetia>(new[]
                    {
                        "ГАЗОСНАБЖЕНИЕ", "газооснаб", "газ","коммуникац"
                    }.Select(s => new SearchCritetia {Text = s}).ToList())
                },
                new TemplateColumn
                {
                    Name = "ВОДОСНАБЖЕНИЕ",
                    CodeName = "SYSTEM_WATER",
                    ColumnIndex = 51,
                    SearchCritetias = new List<SearchCritetia>(new[]
                    {
                        "ВОДОСНАБЖЕНИЕ", "водоснаб", "водос", "вод"
                    }.Select(s => new SearchCritetia {Text = s}).ToList())
                },
                new TemplateColumn
                {
                    Name = "КАНАЛИЗАЦИЯ",
                    CodeName = "SYSTEM_SEWERAGE",
                    ColumnIndex = 52,
                    SearchCritetias = new List<SearchCritetia>(new[]
                    {
                        "КАНАЛИЗАЦИЯ", "канализац", "канализ", "канал"
                    }.Select(s => new SearchCritetia {Text = s}).ToList())
                },
                new TemplateColumn
                {
                    Name = "ЭЛЕКТРОСНАБЖЕНИЕ",
                    CodeName = "SYSTEM_ELECTRICITY",
                    ColumnIndex = 53,
                    SearchCritetias = new List<SearchCritetia>(new[]
                    {
                        "ЭЛЕКТРОСНАБЖЕНИЕ", "электроснаб","элекроснаб", "электрос", "электро", "эле"
                    }.Select(s => new SearchCritetia {Text = s}).ToList())
                },
                new TemplateColumn
                {
                    Name = "ТЕПЛОСНАБЖЕНИЕ",
                    CodeName = "HEAT_SUPPLY",
                    ColumnIndex = 54,
                    SearchCritetias = new List<SearchCritetia>(new[]
                    {
                        "ТЕПЛОСНАБЖЕНИЕ", "теплоснаб", "тепл", "обогр", "отопл"
                    }.Select(s => new SearchCritetia {Text = s}).ToList())
                },
                new TemplateColumn {CodeName = "TRAIN", Name = "Ж/Д_ВЕТКА", ColumnIndex = 55},
                new TemplateColumn {CodeName = "ROAD", Name = "ДОРОГА", ColumnIndex = 56},
                new TemplateColumn {CodeName = "DESCRIPTION", Name = "ОПИСАНИЕ", ColumnIndex = 57,
                    SearchCritetias = new List<SearchCritetia>(new[]
                    {
                        "ОПИСАНИЕ"
                    }.Select(s => new SearchCritetia {Text = s}).ToList())},
                new TemplateColumn {CodeName = "SOURCE_DESC", Name = "ИСТОЧНИК_ИНФОРМАЦИИ",
                    SearchCritetias = new List<SearchCritetia>(new[]
                    {
                        "ИСТОЧНИК_ИНФОРМАЦИИ", "ИСТОЧНИК","информ"
                    }.Select(s => new SearchCritetia {Text = s}).ToList()), ColumnIndex = 58},
                new TemplateColumn {CodeName = "SOURCE_LINK", Name = "ССЫЛКА_НА_ИСТОЧНИК_ИНФОРМАЦИИ", ColumnIndex = 59,
                    SearchCritetias = new List<SearchCritetia>(new[]
                    {
                        "ССЫЛКА_НА_ИСТОЧНИК_ИНФОРМАЦИИ", "ССЫЛКА","URL"
                    }.Select(s => new SearchCritetia {Text = s}).ToList())},
                new TemplateColumn {CodeName = "CONTACTS", Name = "КОНТАКТЫ", ColumnIndex = 60,
                    SearchCritetias = new List<SearchCritetia>(new[]
                    {
                        "ТЕЛЕФОН_ПРОДАВЦА", "КОНТАКТЫ", "ТЕЛЕФОН", "Компания"
                    }.Select(s => new SearchCritetia {Text = s}).ToList())},
                new TemplateColumn {CodeName = "DATE_RESEARCH", Name = "ДАТА_СБОРА_ИНФОРМАЦИИ", ColumnIndex = 61,
                    SearchCritetias = new List<SearchCritetia>(new[]
                    {
                        "ДАТА_РАЗМЕЩЕНИЯ_ИНФОРМАЦИИ", "ДАТА_РАЗМЕЩЕНИЯ", "дата"
                    }.Select(s => new SearchCritetia {Text = s}).ToList())},
                new TemplateColumn
                {
                    Name = "ДАТА ПАРСИНГА",
                    CodeName = "DATE_PARSING",
                    ColumnIndex = 62,
                    SearchCritetias = new List<SearchCritetia>(new[]
                    {
                        "парсинг"
                    }.Select(s => new SearchCritetia {Text = s}).ToList())
                },

            };

#endregion

            var commerceWb = new TemplateWorkbook { WorkbookType = XlTemplateWorkbookType.CommerceProperty };
            commerceWb.Columns.AddRange(columns);

            context.TemplateWorkbooks.Add(commerceWb);

            context.SaveChanges();
        }

        private void InitializeLandWorkbook(TemplateWbsContext context)
        {
            #region Columns
            var columns = new[]
            {
                new TemplateColumn
                {
                    Name = "Уникальный идентификационный номер",
                    CodeName = "OBJECTID",
                    ColumnIndex = 1
                },
                new TemplateColumn
                {
                    Name = "Субъект Российской Федерации",
                    CodeName = "SUBJECT",
                    ColumnIndex = 2,
                    SearchCritetias = new List<SearchCritetia>(new[]
                    {
                        "СУБЪЕКТ_РОССИЙСКОЙ_ФЕДЕРАЦИИ", "субъект", "република", "область", "край","регион"
                    }.Select(s => new SearchCritetia {Text = s}).ToList())
                },
                new TemplateColumn
                {
                    Name = "Муниципальное образование (район)",
                    CodeName = "REGION",
                    ColumnIndex = 3,
                    SearchCritetias = new List<SearchCritetia>(new[]
                    {
                        "МЕСТОПОЛОЖЕНИЕ", "район", "город"
                    }.Select(s => new SearchCritetia {Text = s}).ToList())
                },
                new TemplateColumn {Name = "Поселение", CodeName = "SETTLEMENT", ColumnIndex = 4},
                new TemplateColumn
                {
                    Name = "Ближайший населенный пункт",
                    CodeName = "NEAR_CITY",
                    ColumnIndex = 5,
                    SearchCritetias = new List<SearchCritetia>(new[]
                    {
                        "населенн", "насел","город"
                    }.Select(s => new SearchCritetia {Text = s}).ToList())
                },
                new TemplateColumn
                {
                    Name = "Тип ближайшего населенного пункта",
                    CodeName = "TERRITORY_TYPE",
                    ColumnIndex = 6
                },
                new TemplateColumn
                {
                    Name = "Объект расположен в границах населенного пункта",
                    CodeName = "IN_CITY",
                    ColumnIndex = 7
                },
                new TemplateColumn {Name = "Городской район", CodeName = "VGT", ColumnIndex = 8,
                    SearchCritetias = new List<SearchCritetia>(new[]
                    {
                        "внутригор","территор"
                    }.Select(s => new SearchCritetia {Text = s}).ToList())},
                new TemplateColumn {Name = "Наименование адресного объекта", CodeName = "STREET", ColumnIndex = 9,
                    SearchCritetias = new List<SearchCritetia>(new[]
                    {
                        "улиц"
                    }.Select(s => new SearchCritetia {Text = s}).ToList())},
                new TemplateColumn {Name = "Тип адресного объекта", CodeName = "STREET_TYPE", ColumnIndex = 10},
                new TemplateColumn {Name = "Дом", CodeName = "HOUSE_NUM", ColumnIndex = 11},
                new TemplateColumn {Name = "Литера", CodeName = "LETTER", ColumnIndex = 12},
                new TemplateColumn {Name = "Корпус", CodeName = "BUILDING", ColumnIndex = 13},
                new TemplateColumn {Name = "Строение", CodeName = "STRUCTURE", ColumnIndex = 14},
                new TemplateColumn {Name = "Владение", CodeName = "ESTATE", ColumnIndex = 15},
                new TemplateColumn {Name = "Долгота", CodeName = "LONGITUDE", ColumnIndex = 16},
                new TemplateColumn {Name = "Широта", CodeName = "LATITUDE", ColumnIndex = 17},
                new TemplateColumn {Name = "Трасса", CodeName = "HIGHWAY", ColumnIndex = 18},
                new TemplateColumn
                {
                    Name = "Расстояние до регионального центра",
                    CodeName = "DIST_REG_CENTER",
                    ColumnIndex = 19,
                    SearchCritetias = new List<SearchCritetia>(new[]
                    {
                        "УДАЛЕННОСТЬ", "центр"
                    }.Select(s => new SearchCritetia {Text = s}).ToList())
                },
                new TemplateColumn
                {
                    Name = "Расстояние до ближайшего населенного пункта",
                    CodeName = "DIST_NEAR_CITY",
                    ColumnIndex = 20
                },
                new TemplateColumn
                {
                    Name = "Кадастровый номер земельного участка",
                    CodeName = "CADASTRE_NUM",
                    ColumnIndex = 21
                },
                new TemplateColumn
                {
                    Name = "Предложение (сделка)",
                    CodeName = "OFFER_DEAL",
                    ColumnIndex = 22,
                    SearchCritetias = new List<SearchCritetia>(new[]
                    {
                        "ВИД_СДЕЛКИ", "ВИД СДЕЛКИ", "продажа", "аренда"
                    }.Select(s => new SearchCritetia {Text = s}).ToList())
                },
                new TemplateColumn {Name = "Операция", CodeName = "OPERATION", ColumnIndex = 23,
                    SearchCritetias = new List<SearchCritetia>(new[]
                    {
                        "тип объявл"
                    }.Select(s => new SearchCritetia {Text = s}).ToList())},
                new TemplateColumn
                {
                    Name = "Права на участок",
                    CodeName = "LAW_NOW",
                    ColumnIndex = 24,
                    SearchCritetias = new List<SearchCritetia>(new[]
                    {
                        "ВИД_ПРАВА", "ВИД ПРАВА", "права", "прав","вид собствен"
                    }.Select(s => new SearchCritetia {Text = s}).ToList())
                },
                new TemplateColumn {Name = "Способ реализации", CodeName = "SALE_TYPE", ColumnIndex = 25},
                new TemplateColumn {Name = "Срок аренды", CodeName = "RENTAL_PERIOD", ColumnIndex = 26},
                new TemplateColumn
                {
                    Name = "Цена предложения (сделки)",
                    CodeName = "PRICE",
                    ColumnIndex = 27,
                    SearchCritetias = new List<SearchCritetia>(new[]
                    {
                        "СТОИМОСТЬ", "стоим", "цена", "продаж","стоим","общая стоим"
                    }.Select(s => new SearchCritetia {Text = s}).ToList())
                },
                new TemplateColumn {Name = "Арендная плата", CodeName = "RENT_RATE", ColumnIndex = 28},
                new TemplateColumn
                {
                    Name = "Площадь",
                    CodeName = "AREA_LOT",
                    ColumnIndex = 29,
                    SearchCritetias = new List<SearchCritetia>(new[]
                    {
                        "ПЛОЩАДЬ УЧАСТКА", "ПЛОЩАДЬ_УЧАСТКА", "ПЛОЩАДЬ", "площад", "площ"
                    }.Select(s => new SearchCritetia {Text = s}).ToList())
                },
                new TemplateColumn
                {
                    Name = "Категория земель",
                    CodeName = "LAND_CATEGORY",
                    ColumnIndex = 30,
                    SearchCritetias = new List<SearchCritetia>(new[]
                    {
                        "КАТЕГОРИЯ_ЗЕМЛИ", "категор", "земл"
                    }.Select(s => new SearchCritetia {Text = s}).ToList())
                },
                new TemplateColumn
                {
                    Name = "Вид разрешенного использования",
                    CodeName = "PERMITTED_USE",
                    ColumnIndex = 31
                },
                new TemplateColumn
                {
                    Name = "Вид разрешенного использования текст",
                    CodeName = "PERMITTED_USE_TEXT",
                    ColumnIndex = 32,
                    SearchCritetias = new List<SearchCritetia>(new[]
                    {
                        "ВИД_РАЗРЕШЕННОГО_ИСПОЛЬЗОВАНИЯ", "вид р", "разрешен", "использ"
                    }.Select(s => new SearchCritetia {Text = s}).ToList())
                },
                new TemplateColumn
                {
                    Name = "Газоснабжение",
                    CodeName = "SYSTEM_GAS",
                    ColumnIndex = 33,
                    SearchCritetias = new List<SearchCritetia>(new[]
                    {
                        "ГАЗОСНАБЖЕНИЕ", "газооснаб", "газ","коммуникац"
                    }.Select(s => new SearchCritetia {Text = s}).ToList())
                },
                new TemplateColumn
                {
                    Name = "Водоснабжение",
                    CodeName = "SYSTEM_WATER",
                    ColumnIndex = 34,
                    SearchCritetias = new List<SearchCritetia>(new[]
                    {
                        "ВОДОСНАБЖЕНИЕ", "водоснаб", "водос", "вод"
                    }.Select(s => new SearchCritetia {Text = s}).ToList())
                },
                new TemplateColumn
                {
                    Name = "Канализация",
                    CodeName = "SYSTEM_SEWERAGE",
                    ColumnIndex = 35,
                    SearchCritetias = new List<SearchCritetia>(new[]
                    {
                        "КАНАЛИЗАЦИЯ", "канализац", "канализ", "канал"
                    }.Select(s => new SearchCritetia {Text = s}).ToList())
                },
                new TemplateColumn
                {
                    Name = "Электроснабжение",
                    CodeName = "SYSTEM_ELECTRICITY",
                    ColumnIndex = 36,
                    SearchCritetias = new List<SearchCritetia>(new[]
                    {
                        "ЭЛЕКТРОСНАБЖЕНИЕ", "электроснаб","элекроснаб" ,"электрос", "электро", "элек"
                    }.Select(s => new SearchCritetia {Text = s}).ToList())
                },
                new TemplateColumn
                {
                    Name = "Теплоснабжение",
                    CodeName = "HEAT_SUPPLY",
                    ColumnIndex = 37,
                    SearchCritetias = new List<SearchCritetia>(new[]
                    {
                        "ТЕПЛОСНАБЖЕНИЕ", "теплоснаб", "тепл", "обогр", "отопл"
                    }.Select(s => new SearchCritetia {Text = s}).ToList())
                },
                new TemplateColumn {Name = "Наличие объектов на участке", CodeName = "OBJECT", ColumnIndex = 38,
                    SearchCritetias = new List<SearchCritetia>(new[]
                    {
                        "строения","постройки"
                    }.Select(s => new SearchCritetia {Text = s}).ToList())},
                new TemplateColumn {Name = "Покрытие площадки", CodeName = "SURFACE", ColumnIndex = 39},
                new TemplateColumn {Name = "Дорога", CodeName = "ROAD", ColumnIndex = 40},
                new TemplateColumn
                {
                    Name = "Рельеф",
                    CodeName = "RELIEF",
                    ColumnIndex = 41,
                    SearchCritetias = new List<SearchCritetia>(new[]
                    {
                        "рельеф"
                    }.Select(s => new SearchCritetia {Text = s}).ToList())
                },
                new TemplateColumn
                {
                    Name = "Растительный покров",
                    CodeName = "VEGETATION",
                    ColumnIndex = 42,
                    SearchCritetias = new List<SearchCritetia>(new[]
                    {
                        "растен"
                    }.Select(s => new SearchCritetia {Text = s}).ToList())
                },
                new TemplateColumn
                {
                    Name = "Описание",
                    CodeName = "DESCRIPTION",
                    ColumnIndex = 43,
                    SearchCritetias = new List<SearchCritetia>(new[]
                    {
                        "ОПИСАНИЕ"
                    }.Select(s => new SearchCritetia {Text = s}).ToList())
                },
                new TemplateColumn
                {
                    Name = "Источник информации",
                    CodeName = "SOURCE_DESC",
                    ColumnIndex = 44,
                    SearchCritetias = new List<SearchCritetia>(new[]
                    {
                        "ИСТОЧНИК_ИНФОРМАЦИИ", "ИСТОЧНИК","информ"
                    }.Select(s => new SearchCritetia {Text = s}).ToList())
                },
                new TemplateColumn
                {
                    Name = "Ссылка на источник информации",
                    CodeName = "URL_SALE",
                    ColumnIndex = 45,
                    SearchCritetias = new List<SearchCritetia>(new[]
                    {
                        "ССЫЛКА_НА_ИСТОЧНИК_ИНФОРМАЦИИ", "ССЫЛКА","URL"
                    }.Select(s => new SearchCritetia {Text = s}).ToList())
                },
                new TemplateColumn {Name = "Наименование продавца", CodeName = "SELLER", ColumnIndex = 46,
                    SearchCritetias = new List<SearchCritetia>(new[]
                    {
                        "ТЕЛЕФОН_ПРОДАВЦА", "КОНТАКТЫ", "Компания","продавец"
                    }.Select(s => new SearchCritetia {Text = s}).ToList())},
                new TemplateColumn {Name = "Организационно-правовая форма", CodeName = "OKOPF", ColumnIndex = 47},
                new TemplateColumn {Name = "Адрес сайта в сети интернет", CodeName = "URL_INFO", ColumnIndex = 48},
                new TemplateColumn
                {
                    Name = "Контакты",
                    CodeName = "CONTACTS",
                    ColumnIndex = 49,
                    SearchCritetias = new List<SearchCritetia>(new[]
                    {
                        "ТЕЛЕФОН_ПРОДАВЦА", "КОНТАКТЫ", "ТЕЛ","почт"
                    }.Select(s => new SearchCritetia {Text = s}).ToList())
                },
                new TemplateColumn
                {
                    Name = "Дата размещения информации",
                    CodeName = "DATE_RESEARCH",
                    ColumnIndex = 50,
                    SearchCritetias = new List<SearchCritetia>(new[]
                    {
                        "ДАТА_РАЗМЕЩЕНИЯ_ИНФОРМАЦИИ", "ДАТА_РАЗМЕЩЕНИЯ", "дата"
                    }.Select(s => new SearchCritetia {Text = s}).ToList())
                },
                new TemplateColumn {Name = "Дата отчета по этапу", CodeName = "DATE_IN_BASE", ColumnIndex = 51},
                new TemplateColumn {Name = "Актуальность", CodeName = "ACTUAL", ColumnIndex = 52},
                new TemplateColumn
                {
                    Name = "Дата прозвона",
                    CodeName = "DATE_IS_RINGING",
                    ColumnIndex = 53,
                    SearchCritetias = new List<SearchCritetia>(new[]
                    {
                        "обновлен"
                    }.Select(s => new SearchCritetia {Text = s}).ToList())
                },
                new TemplateColumn {Name = "Результат прозвона", CodeName = "RESULT", ColumnIndex = 54},
                new TemplateColumn
                {
                    Name = "Уточненные (дополненные) характеристики",
                    CodeName = "ADDITIONAL",
                    ColumnIndex = 55
                },
                new TemplateColumn {Name = "Комментарий", CodeName = "COMMENT", ColumnIndex = 56},
                new TemplateColumn {Name = "Товарищества и корпоративы", CodeName = "ASSOCIATIONS", ColumnIndex = 57},
                new TemplateColumn
                {
                    Name = "Дата парсинга",
                    CodeName = "DATE_PARSING",
                    ColumnIndex = 58,
                    SearchCritetias = new List<SearchCritetia>(new[]
                    {
                        "парсинг"
                    }.Select(s => new SearchCritetia {Text = s}).ToList())
                },
                new TemplateColumn {Name = "Ориентир", CodeName = "LAND_MARK", ColumnIndex = 59},
                new TemplateColumn {Name = "Товарищества", CodeName = "SNT", ColumnIndex = 60}
            };
#endregion

            var landWb = new TemplateWorkbook {WorkbookType = XlTemplateWorkbookType.LandProperty};
            landWb.Columns.AddRange(columns);

            context.TemplateWorkbooks.Add(landWb);

            context.SaveChanges();
        }
    }
}