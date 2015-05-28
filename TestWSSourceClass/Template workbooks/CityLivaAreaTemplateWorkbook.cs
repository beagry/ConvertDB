using System.Collections.Generic;
using ExcelRLibrary.TemplateWorkbooks;

namespace Converter.Template_workbooks
{
    internal class CityLivaAreaTemplateWorkbook : TemplateWorkbook
    {
        public CityLivaAreaTemplateWorkbook()
        {
            Columns = new List<JustColumn>
            {
                new JustColumn("OBJECTID", "ПОРЯДКОВЫЙ_НОМЕР", 1),
                new JustColumn("SUBJECT", "СУБЪЕКТ_РОССИЙСКОЙ_ФЕДЕРАЦИИ", 2),
                new JustColumn("REGION", "МУНИЦИПАЛЬНОЕ_ОБРАЗОВАНИЕ_(РАЙОН)", 3),
                new JustColumn("SETTLEMENT", "ПОСЕЛЕНИЕ", 4),
                new JustColumn("CITY", "НАСЕЛЕННЫЙ_ПУНКТ", 5),
                new JustColumn("CITY_TYPE", "ТИП_НАСЕЛЕННОГО_ПУНКТА", 6),
                new JustColumn("VGT", "ВНУТРИГОРОДСКАЯ_ТЕРРИТОРИЯ", 7),
                new JustColumn("STREET", "УЛИЦА", 8),
                new JustColumn("STREET_TYPE", "ТИП_УЛИЦЫ", 9),
                new JustColumn("HOUSE_NUM", "ДОМ", 10),
                new JustColumn("LETTER", "ЛИТЕРА", 11),
                new JustColumn("BUILDING", "КОРПУС", 12),
                new JustColumn("STRUCTURE", "СТРОЕНИЕ", 13),
                new JustColumn("ESTATE", "ВЛАДЕНИЕ", 14),
                new JustColumn("LONGITUDE", "ДОЛГОТА", 15),
                new JustColumn("LATITUDE", "ШИРОТА", 16),
                new JustColumn("METRO", "СТАНЦИЯ_МЕТРО", 17),
                new JustColumn("METRO_DISTMIN", "ДО_МЕТРО_МИНУТ", 18),
                new JustColumn("TRANSPORT", "ПЕШКОМ_ТРАНСПОРТОМ", 19),
                new JustColumn("OBJECT_TYPE", "ТИП_ОБЪЕКТА", 20),
                new JustColumn("OFFER_DEAL", "ПРЕДЛОЖЕНИЕ_СДЕЛКА", 21),
                new JustColumn("ECONOM_CLASS", "ЭКОНОМ", 22),
                new JustColumn("OPERATION", "ОПЕРАЦИЯ", 23),
                new JustColumn("SALE_PRICE", "ЦЕНА _ПРОДАЖИ", 24),
                new JustColumn("RENT_RATE", "АРЕНДНАЯ_ПЛАТА", 25),
                new JustColumn("ROOM_QNT", "КОЛИЧЕСТВО_КОМНАТ", 26),
                new JustColumn("AREA_TOTAL", "ПЛОЩАДЬ_ОБЩАЯ", 27),
                new JustColumn("AREA_LIVING", "ПЛОЩАДЬ_ЖИЛАЯ", 28),
                new JustColumn("AREA_KITCHEN", "ПЛОЩАДЬ_КУХНИ", 29),
                new JustColumn("FLOOR_NUM", "ЭТАЖ", 30),
                new JustColumn("FLOOR_QNT", "ЭТАЖНОСТЬ", 31),
                new JustColumn("MATERIAL_WALL", "МАТЕРИАЛ_СТЕН", 32),
                new JustColumn("YEAR_BUILT", "ГОД_ПОСТРОЙКИ", 33),
                new JustColumn("ROOM_TYPE", "РАСПОЛОЖЕНИЕ_КОМНАТ", 34),
                new JustColumn("BALCONY", "ЛОДЖИЯ_БАЛКОН", 35),
                new JustColumn("BATHROOM", "САНУЗЕЛ", 36),
                new JustColumn("WINDOWS", "ОКНА", 37),
                new JustColumn("CONDITION", "СОСТОЯНИЕ", 38),
                new JustColumn("CONSIERGE", "КОНСЬЕРЖ", 39),
                new JustColumn("DESCRIPTION", "ОПИСАНИЕ", 40),
                new JustColumn("SOURCE_DESC", "ИСТОЧНИК_ИНФОРМАЦИИ", 41),
                new JustColumn("SOURCE_LINK", "ССЫЛКА_НА_ИСТОЧНИК_ИНФОРМАЦИИ", 42),
                new JustColumn("CONTACTS", "КОНТАКТЫ", 43),
                new JustColumn("DATE_RESEARCH", "ДАТА_РАЗМЕЩЕНИЯ_ИНФОРМАЦИИ", 44),
                new JustColumn("DATE_PARSING", "ДАТА_ПАРСИНГА", 45)
            };
        }
    }
}