using System.ComponentModel;

namespace Converter.Template_workbooks
{
    /// <summary>
    ///     Перечень шаблонных книг
    /// </summary>
    public enum XlTemplateWorkbookType
    {
        [Description("Земельные участки")] LandProperty, //Земельные участки
        [Description("Коммерция")] CommerceProperty //Коммерческая нд
//        [Description("Загородка")]
//        CountyLiveArea,//Загородка
//        [Description("Городское жильё")]
//        CityLivaArea//Городское жильё
    }

    public static class TemplateEnumExtention
    {
        public static TemplateWorkbook GetWorkbook(this XlTemplateWorkbookType xlTemplate)
        {
            switch (xlTemplate)
            {
//                    case XlTemplateWorkbookTypes.CityLivaArea:
//                    return new CityLivaAreaTemplateWorkbook();
                case XlTemplateWorkbookType.CommerceProperty:
                    return new CommercePropertyTemplateWorkbook();
//                    case XlTemplateWorkbookTypes.CountyLiveArea:
//                    return new CountryLiveAreaTemplateWorkbook();
                case XlTemplateWorkbookType.LandProperty:
                    return new LandPropertyTemplateWorkbook();
            }
            return null;
        }
    }
}