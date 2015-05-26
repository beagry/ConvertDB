using System.ComponentModel;
using Converter.Template_workbooks;

namespace Converter
{
    /// <summary>
    /// Перечень шаблонных книг
    /// </summary>
    public enum XlTemplateWorkbookTypes
    {
        [Description("Земельные участки")]
        LandProperty, //Земельные участки
        [Description("Коммерция")]
        CommerceProperty, //Коммерческая нд
        [Description("Загородка")]
        CountyLiveArea,//Загородка
        [Description("Городское жильё")]
        CityLivaArea//Городское жильё
    }

    public static class TemplateEnumExtention
    {
        public static TemplateWorkbook GetWorkbook(this XlTemplateWorkbookTypes xlTemplate)
        {
            switch (xlTemplate)
            {
                    case XlTemplateWorkbookTypes.CityLivaArea:
                    return new CityLivaAreaTemplateWorkbook();
                    case XlTemplateWorkbookTypes.CommerceProperty:
                    return new CommercePropertyTemplateWorkbook();
                    case XlTemplateWorkbookTypes.CountyLiveArea:
                    return new CountryLiveAreaTemplateWorkbook();
                    case XlTemplateWorkbookTypes.LandProperty:
                    return new LandPropertyTemplateWorkbook();
            }
            return null;
        }
    }
}