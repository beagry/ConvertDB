using System.Collections.Generic;
using ExcelRLibrary.TemplateWorkbooks;

namespace Converter.Template_workbooks
{
    class CountryLiveAreaTemplateWorkbook:TemplateWorkbook
    {
        public CountryLiveAreaTemplateWorkbook()
        {
            Columns = new List<JustColumn>()
            {
                new JustColumn("OBJECTID","ПОРЯДКОВЫЙ_НОМЕР",1),
                new JustColumn("SUBJECT","СУБЪЕКТ_РОССИЙСКОЙ_ФЕДЕРАЦИИ",2),
                new JustColumn("REGION","МУНИЦИПАЛЬНОЕ_ОБРАЗОВАНИЕ_(РАЙОН)",3),
                new JustColumn("SETTLEMENT","ПОСЕЛЕНИЕ",4),
                new JustColumn("CITY","НАСЕЛЕННЫЙ_ПУНКТ",5),
                new JustColumn("CITY_TYPE","ТИП_НАСЕЛЕННОГО_ПУНКТА",6),
                new JustColumn("IN_CITY","ОБЪЕКТ_РАСПОЛОЖЕН_В ГРАНИЦАХ_НАСЕЛЕННОГО_ПУНКТА",7),
                new JustColumn("VGT","ВНУТРИГОРОДСКАЯ_ТЕРРИТОРИЯ",8),
                new JustColumn("STREET","УЛИЦА",9),
                new JustColumn("STREET_TYPE","ТИП_УЛИЦЫ",10),
                new JustColumn("HOUSE_NUM","ДОМ",11),
                new JustColumn("LETTER","ЛИТЕРА",12),
                new JustColumn("BUILDING","КОРПУС",13),
                new JustColumn("STRUCTURE","СТРОЕНИЕ",14),
                new JustColumn("ESTATE","ВЛАДЕНИЕ",15),
                new JustColumn("LONGITUDE","ДОЛГОТА",16),
                new JustColumn("LATITUDE","ШИРОТА",17),
                new JustColumn("HIGHWAY","ТРАССА",18),
                new JustColumn("DIST_REG_CENTER","УДАЛЕННОСТЬ_ОТ_РЕГИОНАЛЬНОГО_ЦЕНТРА",19),
                new JustColumn("DIST_NEAR_CITY","РАССТОЯНИЕ_ОТ_БЛИЖАЙШЕГО_НАСЕЛЕННОГО_ПУНКТА",20),
                new JustColumn("CADASTRE_NUM","КАДАСТРОВЫЙ_НОМЕР_ЗЕМЕЛЬНОГО_УЧАСТКА",21),
                new JustColumn("OBJECT_TYPE","ТИП_ОБЪЕКТА",22),
                new JustColumn("OFFER_DEAL","ПРЕДЛОЖЕНИЕ_СДЕЛКА",23),
                new JustColumn("OPERATION","ОПЕРАЦИЯ",24),
                new JustColumn("SALE_PRICE","ЦЕНА_ПРОДАЖИ",25),
                new JustColumn("RENT_RATE","АРЕНДНАЯ_ПЛАТА",26),
                new JustColumn("AREA_TOTAL","ПЛОЩАДЬ_ОБЩАЯ",27),
                new JustColumn("ROOM_QNT","КОЛИЧЕСТВО_КОМНАТ",28),
                new JustColumn("FLOOR_QNT","ЭТАЖНОСТЬ",29),
                new JustColumn("MATERIAL_WALL","МАТЕРИАЛ_СТЕН",30),
                new JustColumn("YEAR_BUILD","ГОД_ПОСТРОЙКИ",31),
                new JustColumn("AREA_LOT","ПЛОЩАДЬ_УЧАСТКА",32),
                new JustColumn("OBJECT","ДОПОЛНИТЕЛЬНЫЕ_ПОСТРОЙКИ",33),
                new JustColumn("SYSTEM_GAS","ГАЗОСНАБЖЕНИЕ",34),
                new JustColumn("SYSTEM_WATER","ВОДОСНАБЖЕНИЕ",35),
                new JustColumn("SYSTEM_SEWERAGE","КАНАЛИЗАЦИЯ",36),
                new JustColumn("SYSTEM_ELECTRICITY","ЭЛЕКТРОСНАБЖЕНИЕ",37),
                new JustColumn("HEAT_SUPPLY","ТЕПЛОСНАБЖЕНИЕ",38),
                new JustColumn("FOREST","ЛЕС",39),
                new JustColumn("WATER","ВОДОЕМ",40),
                new JustColumn("SECURITY","БЕЗОПАСНОСТЬ",41),
                new JustColumn("DESCRIPTION","ОПИСАНИЕ",42),
                new JustColumn("SOURCE_DESC","ИСТОЧНИК_ИНФОРМАЦИИ",43),
                new JustColumn("SOURCE_LINK","ССЫЛКА_НА_ИСТОЧНИК_ИНФОРМАЦИИ",44),
                new JustColumn("CONTACTS","КОНТАКТЫ",45),
                new JustColumn("DATE_RESEARCH","ДАТА_РАЗМЕЩЕНИЯ_ИНФОРМАЦИИ",46),
                new JustColumn("DATE_PARSING","ДАТА_ПАРСИНГА",47),

            };
        }
    }
}
