using System.ComponentModel;
using Converter.Template_workbooks;

namespace Converter
{
    /// <summary>
    /// �������� ��������� ����
    /// </summary>
    public enum XlTemplateWorkbookTypes
    {
        [Description("��������� �������")]
        LandProperty, //��������� �������
        [Description("���������")]
        CommerceProperty, //������������ ��
        [Description("���������")]
        CountyLiveArea,//���������
        [Description("��������� �����")]
        CityLivaArea//��������� �����
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