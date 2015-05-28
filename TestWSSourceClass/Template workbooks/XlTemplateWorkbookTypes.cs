using System.ComponentModel;

namespace Converter.Template_workbooks
{
    /// <summary>
    ///     �������� ��������� ����
    /// </summary>
    public enum XlTemplateWorkbookType
    {
        [Description("��������� �������")] LandProperty, //��������� �������
        [Description("���������")] CommerceProperty //������������ ��
//        [Description("���������")]
//        CountyLiveArea,//���������
//        [Description("��������� �����")]
//        CityLivaArea//��������� �����
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