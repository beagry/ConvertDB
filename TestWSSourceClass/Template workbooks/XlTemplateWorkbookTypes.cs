using System.ComponentModel;

namespace Converter.Template_workbooks
{
    /// <summary>
    ///     Перечень шаблонных книг
    /// </summary>
    public enum XlTemplateWorkbookType
    {
        [Description("Земельные участки")] 
        LandProperty,

        [Description("Коммерция")] 
        CommerceProperty
//        [Description("Загородка")]
//        CountyLiveArea,
//        [Description("Городское жильё")]
//        CityLivaArea
    }
}