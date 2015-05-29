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
        }

        private void InitializeLandWorkbook(TemplateWbsContext context)
        {

            var LandPlusCommerceColumns = 

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
                        "СУБЪЕКТ_РОССИЙСКОЙ_ФЕДЕРАЦИИ", "субъект", "република", "область", "край"
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
                        "населенн", "насел"
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
                new TemplateColumn {Name = "Городской район", CodeName = "VGT", ColumnIndex = 8},
                new TemplateColumn {Name = "Наименование адресного объекта", CodeName = "STREET", ColumnIndex = 9},
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
                new TemplateColumn {Name = "Операция", CodeName = "OPERATION", ColumnIndex = 23},
                new TemplateColumn
                {
                    Name = "Права на участок",
                    CodeName = "LAW_NOW",
                    ColumnIndex = 24,
                    SearchCritetias = new List<SearchCritetia>(new[]
                    {
                        "ВИД_ПРАВА", "ВИД ПРАВА", "права", "прав"
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
                        "СТОИМОСТЬ", "стоим", "цена", "продаж"
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
                        "ГАЗОСНАБЖЕНИЕ", "газооснаб", "газ"
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
                        "ЭЛЕКТРОСНАБЖЕНИЕ", "электроснаб", "электрос", "электро", "эле"
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
                new TemplateColumn {Name = "Наличие объектов на участке", CodeName = "OBJECT", ColumnIndex = 38},
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
                        "ИСТОЧНИК_ИНФОРМАЦИИ", "ИСТОЧНИК"
                    }.Select(s => new SearchCritetia {Text = s}).ToList())
                },
                new TemplateColumn
                {
                    Name = "Ссылка на источник информации",
                    CodeName = "URL_SALE",
                    ColumnIndex = 45,
                    SearchCritetias = new List<SearchCritetia>(new[]
                    {
                        "ССЫЛКА_НА_ИСТОЧНИК_ИНФОРМАЦИИ", "ССЫЛКА"
                    }.Select(s => new SearchCritetia {Text = s}).ToList())
                },
                new TemplateColumn {Name = "Наименование продавца", CodeName = "SELLER", ColumnIndex = 46},
                new TemplateColumn {Name = "Организационно-правовая форма", CodeName = "OKOPF", ColumnIndex = 47},
                new TemplateColumn {Name = "Адрес сайта в сети интернет", CodeName = "URL_INFO", ColumnIndex = 48},
                new TemplateColumn
                {
                    Name = "Телефон продавца",
                    CodeName = "CONTACTS",
                    ColumnIndex = 49,
                    SearchCritetias = new List<SearchCritetia>(new[]
                    {
                        "ТЕЛЕФОН_ПРОДАВЦА", "КОНТАКТЫ", "ТЕЛЕФОН", "Компания"
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
                new TemplateColumn {Name = "Ориентиры", CodeName = "LAND_MARK", ColumnIndex = 59},
                new TemplateColumn {Name = "Товарищества", CodeName = "SNT", ColumnIndex = 60}
            };

            var landWb = new TemplateWorkbook {WorkbookType = XlTemplateWorkbookType.LandProperty};
            landWb.Columns.AddRange(columns);

            context.TemplateWorkbooks.Add(landWb);

            context.SaveChanges();
        }
    }
}