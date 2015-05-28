using System.Collections.Generic;
using ExcelRLibrary.TemplateWorkbooks;

namespace Converter.Template_workbooks
{
    class CommercePropertyTemplateWorkbook:TemplateWorkbook
    {
        public CommercePropertyTemplateWorkbook()
        {
            Columns = new List<JustColumn> 
            #region Columns Initialize
            {
                new JustColumn("ID","ПОРЯДКОВЫЙ_НОМЕР",1),
                new JustColumn("SUBJECT","СУБЪЕКТ_РОССИЙСКОЙ_ФЕДЕРАЦИИ",2),
                new JustColumn("REGION","МУНИЦИПАЛЬНОЕ_ОБРАЗОВАНИЕ_(РАЙОН)",3),
                new JustColumn("SETTLEMENT","ПОСЕЛЕНИЕ",4),
                new JustColumn("CITY","НАСЕЛЕННЫЙ_ПУНКТ",5),
                new JustColumn("CITY_TYPE","ТИП_НАСЕЛЕННОГО_ПУНКТА",6),
                new JustColumn("VGT","ВНУТРИГОРОДСКАЯ_ТЕРРИТОРИЯ",7),
                new JustColumn("STREET","УЛИЦА",8),
                new JustColumn("STREET_TYPE","ТИП_УЛИЦЫ",9),
                new JustColumn("HOUSE_NUM","ДОМ",10),
                new JustColumn("LETTER","ЛИТЕРА",11),
                new JustColumn("BUILDING","КОРПУС",12),
                new JustColumn("STRUCTURE","СТРОЕНИЕ",13),
                new JustColumn("ESTATE","ВЛАДЕНИЕ",14),
                new JustColumn("LONGITUDE","ДОЛГОТА",15),
                new JustColumn("LATITUDE","ШИРОТА",16),
                new JustColumn("DIST_REG_CENTER","УДАЛЕННОСТЬ_ОТ_РЕГИОНАЛЬНОГО_ЦЕНТРА",17),
                new JustColumn("CADASTRE_NUM","КАДАСТРОВЫЙ_НОМЕР_ЗЕМЕЛЬНОГО_УЧАСТКА",18),
                new JustColumn("METRO","СТАНЦИЯ_МЕТРО",19),
                new JustColumn("METRO_DISTMIN","ДО_МЕТРО_МИНУТ",20),
                new JustColumn("TRANSPORT","ПЕШКОМ_ТРАНСПОРТОМ",21),
                new JustColumn("SEGMENT","СЕГМЕНТ",22),
                new JustColumn("BUILDING_TYPE","ТИП_ПОСТРОЙКИ",23),
                new JustColumn("CENTER_NAME","НАИМЕНОВАНИЕ_ЦЕНТРА",24),
                new JustColumn("OBJECT_TYPE","ТИП_ОБЪЕКТА",25),
                new JustColumn("OBJECT_PURPOSE","НАЗНАЧЕНИЕ_ОБЪЕКТА",26),
                new JustColumn("CLASS_TYPE","ПОТРЕБИТЕЛЬСКИЙ_КЛАСС",27),
                new JustColumn("OPERATION","ОПЕРАЦИЯ",28),
                new JustColumn("SALE_PRICE","ЦЕНА _ПРОДАЖИ",29),
                new JustColumn("RENT_RATE","АРЕНДНАЯ_ПЛАТА",30),
                new JustColumn("AREA","ПЛОЩАДЬ",31),
                new JustColumn("PRICE_FOR_UNIT","ЦЕНА_ЗА_М2",32),
                new JustColumn("OPERATING_COSTS","ЭКСПЛУАТАЦИОННЫЕ_РАСХОДЫ",33),
                new JustColumn("FLOOR","ЭТАЖ",34),
                new JustColumn("FLOOR_QNT_MIN","ЭТАЖНОСТЬ_МИНИМАЛЬНАЯ",35),
                new JustColumn("FLOOR_QNT_MAX","ЭТАЖНОСТЬ_МАКСИМАЛЬНАЯ",36),
                new JustColumn("YEAR_BUILD","ГОД_ПОСТРОЙКИ",37),
                new JustColumn("MATERIAL_WALL","МАТЕРИАЛ_СТЕН",38),
                new JustColumn("HEIGHT_FLOOR","ВЫСОТА_ПОТОЛКА",39),
                new JustColumn("COLUMN_DIST","ШАГ_КОЛОНН",40),
                new JustColumn("LAYOUT","ПЛАНИРОВКА",41),
                new JustColumn("ROOM_QNT","КОЛИЧЕСТВО_КОМНАТ",42),
                new JustColumn("AREA_TOTAL","ПЛОЩАДЬ_ОБЩАЯ",43),
                new JustColumn("AREA_LOT","ПЛОЩАДЬ_ЗЕМЕЛЬНОГО_УЧАСТКА_ОБЪЕКТА",44),
                new JustColumn("CONDITION","СОСТОЯНИЕ",45),
                new JustColumn("SECURITY","БЕЗОПАСНОСТЬ",46),
                new JustColumn("FLOOR_LOAD","ДОПУСТИМАЯ НАГРУЗКА НА ПОЛ",47),
                new JustColumn("CONDITIONING","КОНДИЦИОНИРОВАНИЕ",48),
                new JustColumn("VENT","ВЕНТИЛЯЦИЯ",49),
                new JustColumn("SYSTEM_GAS","ГАЗОСНАБЖЕНИЕ",50),
                new JustColumn("SYSTEM_WATER","ВОДОСНАБЖЕНИЕ",51),
                new JustColumn("SYSTEM_SEWERAGE","КАНАЛИЗАЦИЯ",52),
                new JustColumn("SYSTEM_ELECTRICITY","ЭЛЕКТРОСНАБЖЕНИЕ",53),
                new JustColumn("HEAT_SUPPLY","ТЕПЛОСНАБЖЕНИЕ",54),
                new JustColumn("TRAIN","Ж/Д_ВЕТКА",55),
                new JustColumn("ROAD","ДОРОГА",56),
                new JustColumn("DESCRIPTION","ОПИСАНИЕ",57),
                new JustColumn("SOURCE_DESC","ИСТОЧНИК_ИНФОРМАЦИИ",58),
                new JustColumn("SOURCE_LINK","ССЫЛКА_НА_ИСТОЧНИК_ИНФОРМАЦИИ",59),
                new JustColumn("CONTACTS","КОНТАКТЫ",60),
                new JustColumn("DATE_RESEARCH","ДАТА_СБОРА_ИНФОРМАЦИИ",61),
                new JustColumn("LANDMARK", "ОРИЕНТИР", 62),
                new JustColumn("DATE_PARSING", "ДАТА_ПАРСИНГА", 63),
                #endregion

            };
        }
    }
}
