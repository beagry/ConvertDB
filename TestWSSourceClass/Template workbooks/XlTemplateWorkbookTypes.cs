using System.ComponentModel;

namespace Converter.Template_workbooks
{
    /// <summary>
    ///     �������� ��������� ����
    /// </summary>
    public enum XlTemplateWorkbookType
    {
        [Description("��������� �������")] 
        LandProperty,

        [Description("���������")] 
        CommerceProperty
//        [Description("���������")]
//        CountyLiveArea,
//        [Description("��������� �����")]
//        CityLivaArea
    }
}