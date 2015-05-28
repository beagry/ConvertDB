using System.Collections.Generic;
using ExcelRLibrary.TemplateWorkbooks;

namespace Converter.Template_workbooks
{
    public class LandPropertyTemplateWorkbook : TemplateWorkbook
    {
        public LandPropertyTemplateWorkbook()
        {
            Columns = new List<JustColumn>
                #region Columns Initialize

            {
                new JustColumn("OBJECTID", "Уникальный идентификационный номер", 1),
                new JustColumn("SUBJECT", "Субъект Российской Федерации", 2),
                new JustColumn("REGION", "Муниципальное образование (район)", 3),
                new JustColumn("SETTLEMENT", "Поселение", 4),
                new JustColumn("NEAR_CITY", "Ближайший населенный пункт", 5),
                new JustColumn("TERRITORY_TYPE", "Тип ближайшего населенного пункта", 6),
                new JustColumn("IN_CITY", "Объект расположен в границах населенного пункта", 7),
                new JustColumn("VGT", "Городской район", 8),
                new JustColumn("STREET", "Наименование адресного объекта", 9),
                new JustColumn("STREET_TYPE", "Тип адресного объекта", 10),
                new JustColumn("HOUSE_NUM", "Дом", 11),
                new JustColumn("LETTER", "Литера", 12),
                new JustColumn("BUILDING", "Корпус", 13),
                new JustColumn("STRUCTURE", "Строение", 14),
                new JustColumn("ESTATE", "Владение", 15),
                new JustColumn("LONGITUDE", "Долгота", 16),
                new JustColumn("LATITUDE", "Широта", 17),
                new JustColumn("HIGHWAY", "Трасса", 18),
                new JustColumn("DIST_REG_CENTER", "Расстояние до регионального центра", 19),
                new JustColumn("DIST_NEAR_CITY", "Расстояние до ближайшего населенного пункта", 20),
                new JustColumn("CADASTRE_NUM", "Кадастровый номер земельного участка", 21),
                new JustColumn("OFFER_DEAL", "Предложение (сделка)", 22),
                new JustColumn("OPERATION", "Операция", 23),
                new JustColumn("LAW_NOW", "Права на участок", 24),
                new JustColumn("SALE_TYPE", "Способ реализации", 25),
                new JustColumn("RENTAL_PERIOD", "Срок аренды", 26),
                new JustColumn("PRICE", "Цена предложения (сделки)", 27),
                new JustColumn("RENT_RATE", "Арендная плата", 28),
                new JustColumn("AREA_LOT", "Площадь", 29),
                new JustColumn("LAND_CATEGORY", "Категория земель", 30),
                new JustColumn("PERMITTED_USE", "Вид разрешенного использования", 31),
                new JustColumn("PERMITTED_USE_TEXT", "Вид разрешенного использования текст", 32),
                new JustColumn("SYSTEM_GAS", "Газоснабжение", 33),
                new JustColumn("SYSTEM_WATER", "Водоснабжение", 34),
                new JustColumn("SYSTEM_SEWERAGE", "Канализация", 35),
                new JustColumn("SYSTEM_ELECTRICITY", "Электроснабжение", 36),
                new JustColumn("HEAT_SUPPLY", "Теплоснабжение", 37),
                new JustColumn("OBJECT", "Наличие объектов на участке", 38),
                new JustColumn("SURFACE", "Покрытие площадки", 39),
                new JustColumn("ROAD", "Дорога", 40),
                new JustColumn("RELIEF", "Рельеф", 41),
                new JustColumn("VEGETATION", "Растительный покров", 42),
                new JustColumn("DESCRIPTION", "Описание", 43),
                new JustColumn("SOURCE_DESC", "Источник информации", 44),
                new JustColumn("URL_SALE", "Ссылка на источник информации", 45),
                new JustColumn("SELLER", "Наименование продавца", 46),
                new JustColumn("OKOPF", "Организационно-правовая форма", 47),
                new JustColumn("URL_INFO", "Адрес сайта в сети интернет", 48),
                new JustColumn("CONTACTS", "Контакты", 49),
                new JustColumn("DATE_RESEARCH", "ДАТА_РАЗМЕЩЕНИЯ_ИНФОРМАЦИИ", 50),
                new JustColumn("DATE_IN_BASE", "Дата отчета по этапу", 51),
                new JustColumn("ACTUAL", "Актуальность", 52),
                new JustColumn("DATE_IS_RINGING", "Дата прозвона", 53),
                new JustColumn("RESULT", "Результат прозвона", 54),
                new JustColumn("COMMENT", "Комментарий", 55),
                new JustColumn("ADDITIONAL", "Уточненные (дополненные) характеристики", 56),
                new JustColumn("ASSOCIATIONS", "Товарищества и корпоративы", 57),
                new JustColumn("DATE_PARSING", "ДАТА_ПАРСИНГА", 58),
                new JustColumn("LAND_MARK", "Ориентиры", 59),
                new JustColumn("SNT", "Товарищества", 60)

                #endregion
            };
        }
    }
}