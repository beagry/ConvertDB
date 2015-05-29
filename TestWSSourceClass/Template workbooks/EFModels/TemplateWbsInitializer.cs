using System.Collections.Generic;
using System.Data.Entity;
using System.Linq;

namespace Converter.Template_workbooks.EFModels
{
    internal class TemplateWbsInitializer : DropCreateDatabaseIfModelChanges<TemplateWbsContext>
    {
        protected override void Seed(TemplateWbsContext context)
        {
            base.Seed(context);

            InitializeLandWorkbook(context);
        }

        private void InitializeLandWorkbook(TemplateWbsContext context)
        {

            var LandPlusCommerceColumns = 

            var columns = new[]
            {
                new TemplateColumn
                {
                    Name = "���������� ����������������� �����",
                    CodeName = "OBJECTID",
                    ColumnIndex = 1
                },
                new TemplateColumn
                {
                    Name = "������� ���������� ���������",
                    CodeName = "SUBJECT",
                    ColumnIndex = 2,
                    SearchCritetias = new List<SearchCritetia>(new[]
                    {
                        "�������_����������_���������", "�������", "���������", "�������", "����"
                    }.Select(s => new SearchCritetia {Text = s}).ToList())
                },
                new TemplateColumn
                {
                    Name = "������������� ����������� (�����)",
                    CodeName = "REGION",
                    ColumnIndex = 3,
                    SearchCritetias = new List<SearchCritetia>(new[]
                    {
                        "��������������", "�����", "�����"
                    }.Select(s => new SearchCritetia {Text = s}).ToList())
                },
                new TemplateColumn {Name = "���������", CodeName = "SETTLEMENT", ColumnIndex = 4},
                new TemplateColumn
                {
                    Name = "��������� ���������� �����",
                    CodeName = "NEAR_CITY",
                    ColumnIndex = 5,
                    SearchCritetias = new List<SearchCritetia>(new[]
                    {
                        "��������", "�����"
                    }.Select(s => new SearchCritetia {Text = s}).ToList())
                },
                new TemplateColumn
                {
                    Name = "��� ���������� ����������� ������",
                    CodeName = "TERRITORY_TYPE",
                    ColumnIndex = 6
                },
                new TemplateColumn
                {
                    Name = "������ ���������� � �������� ����������� ������",
                    CodeName = "IN_CITY",
                    ColumnIndex = 7
                },
                new TemplateColumn {Name = "��������� �����", CodeName = "VGT", ColumnIndex = 8},
                new TemplateColumn {Name = "������������ ��������� �������", CodeName = "STREET", ColumnIndex = 9},
                new TemplateColumn {Name = "��� ��������� �������", CodeName = "STREET_TYPE", ColumnIndex = 10},
                new TemplateColumn {Name = "���", CodeName = "HOUSE_NUM", ColumnIndex = 11},
                new TemplateColumn {Name = "������", CodeName = "LETTER", ColumnIndex = 12},
                new TemplateColumn {Name = "������", CodeName = "BUILDING", ColumnIndex = 13},
                new TemplateColumn {Name = "��������", CodeName = "STRUCTURE", ColumnIndex = 14},
                new TemplateColumn {Name = "��������", CodeName = "ESTATE", ColumnIndex = 15},
                new TemplateColumn {Name = "�������", CodeName = "LONGITUDE", ColumnIndex = 16},
                new TemplateColumn {Name = "������", CodeName = "LATITUDE", ColumnIndex = 17},
                new TemplateColumn {Name = "������", CodeName = "HIGHWAY", ColumnIndex = 18},
                new TemplateColumn
                {
                    Name = "���������� �� ������������� ������",
                    CodeName = "DIST_REG_CENTER",
                    ColumnIndex = 19,
                    SearchCritetias = new List<SearchCritetia>(new[]
                    {
                        "�����������", "�����"
                    }.Select(s => new SearchCritetia {Text = s}).ToList())
                },
                new TemplateColumn
                {
                    Name = "���������� �� ���������� ����������� ������",
                    CodeName = "DIST_NEAR_CITY",
                    ColumnIndex = 20
                },
                new TemplateColumn
                {
                    Name = "����������� ����� ���������� �������",
                    CodeName = "CADASTRE_NUM",
                    ColumnIndex = 21
                },
                new TemplateColumn
                {
                    Name = "����������� (������)",
                    CodeName = "OFFER_DEAL",
                    ColumnIndex = 22,
                    SearchCritetias = new List<SearchCritetia>(new[]
                    {
                        "���_������", "��� ������", "�������", "������"
                    }.Select(s => new SearchCritetia {Text = s}).ToList())
                },
                new TemplateColumn {Name = "��������", CodeName = "OPERATION", ColumnIndex = 23},
                new TemplateColumn
                {
                    Name = "����� �� �������",
                    CodeName = "LAW_NOW",
                    ColumnIndex = 24,
                    SearchCritetias = new List<SearchCritetia>(new[]
                    {
                        "���_�����", "��� �����", "�����", "����"
                    }.Select(s => new SearchCritetia {Text = s}).ToList())
                },
                new TemplateColumn {Name = "������ ����������", CodeName = "SALE_TYPE", ColumnIndex = 25},
                new TemplateColumn {Name = "���� ������", CodeName = "RENTAL_PERIOD", ColumnIndex = 26},
                new TemplateColumn
                {
                    Name = "���� ����������� (������)",
                    CodeName = "PRICE",
                    ColumnIndex = 27,
                    SearchCritetias = new List<SearchCritetia>(new[]
                    {
                        "���������", "�����", "����", "������"
                    }.Select(s => new SearchCritetia {Text = s}).ToList())
                },
                new TemplateColumn {Name = "�������� �����", CodeName = "RENT_RATE", ColumnIndex = 28},
                new TemplateColumn
                {
                    Name = "�������",
                    CodeName = "AREA_LOT",
                    ColumnIndex = 29,
                    SearchCritetias = new List<SearchCritetia>(new[]
                    {
                        "������� �������", "�������_�������", "�������", "������", "����"
                    }.Select(s => new SearchCritetia {Text = s}).ToList())
                },
                new TemplateColumn
                {
                    Name = "��������� ������",
                    CodeName = "LAND_CATEGORY",
                    ColumnIndex = 30,
                    SearchCritetias = new List<SearchCritetia>(new[]
                    {
                        "���������_�����", "�������", "����"
                    }.Select(s => new SearchCritetia {Text = s}).ToList())
                },
                new TemplateColumn
                {
                    Name = "��� ������������ �������������",
                    CodeName = "PERMITTED_USE",
                    ColumnIndex = 31
                },
                new TemplateColumn
                {
                    Name = "��� ������������ ������������� �����",
                    CodeName = "PERMITTED_USE_TEXT",
                    ColumnIndex = 32,
                    SearchCritetias = new List<SearchCritetia>(new[]
                    {
                        "���_������������_�������������", "��� �", "��������", "�������"
                    }.Select(s => new SearchCritetia {Text = s}).ToList())
                },
                new TemplateColumn
                {
                    Name = "�������������",
                    CodeName = "SYSTEM_GAS",
                    ColumnIndex = 33,
                    SearchCritetias = new List<SearchCritetia>(new[]
                    {
                        "�������������", "���������", "���"
                    }.Select(s => new SearchCritetia {Text = s}).ToList())
                },
                new TemplateColumn
                {
                    Name = "�������������",
                    CodeName = "SYSTEM_WATER",
                    ColumnIndex = 34,
                    SearchCritetias = new List<SearchCritetia>(new[]
                    {
                        "�������������", "��������", "�����", "���"
                    }.Select(s => new SearchCritetia {Text = s}).ToList())
                },
                new TemplateColumn
                {
                    Name = "�����������",
                    CodeName = "SYSTEM_SEWERAGE",
                    ColumnIndex = 35,
                    SearchCritetias = new List<SearchCritetia>(new[]
                    {
                        "�����������", "���������", "�������", "�����"
                    }.Select(s => new SearchCritetia {Text = s}).ToList())
                },
                new TemplateColumn
                {
                    Name = "����������������",
                    CodeName = "SYSTEM_ELECTRICITY",
                    ColumnIndex = 36,
                    SearchCritetias = new List<SearchCritetia>(new[]
                    {
                        "����������������", "�����������", "��������", "�������", "���"
                    }.Select(s => new SearchCritetia {Text = s}).ToList())
                },
                new TemplateColumn
                {
                    Name = "��������������",
                    CodeName = "HEAT_SUPPLY",
                    ColumnIndex = 37,
                    SearchCritetias = new List<SearchCritetia>(new[]
                    {
                        "��������������", "���������", "����", "�����", "�����"
                    }.Select(s => new SearchCritetia {Text = s}).ToList())
                },
                new TemplateColumn {Name = "������� �������� �� �������", CodeName = "OBJECT", ColumnIndex = 38},
                new TemplateColumn {Name = "�������� ��������", CodeName = "SURFACE", ColumnIndex = 39},
                new TemplateColumn {Name = "������", CodeName = "ROAD", ColumnIndex = 40},
                new TemplateColumn
                {
                    Name = "������",
                    CodeName = "RELIEF",
                    ColumnIndex = 41,
                    SearchCritetias = new List<SearchCritetia>(new[]
                    {
                        "������"
                    }.Select(s => new SearchCritetia {Text = s}).ToList())
                },
                new TemplateColumn
                {
                    Name = "������������ ������",
                    CodeName = "VEGETATION",
                    ColumnIndex = 42,
                    SearchCritetias = new List<SearchCritetia>(new[]
                    {
                        "������"
                    }.Select(s => new SearchCritetia {Text = s}).ToList())
                },
                new TemplateColumn
                {
                    Name = "��������",
                    CodeName = "DESCRIPTION",
                    ColumnIndex = 43,
                    SearchCritetias = new List<SearchCritetia>(new[]
                    {
                        "��������"
                    }.Select(s => new SearchCritetia {Text = s}).ToList())
                },
                new TemplateColumn
                {
                    Name = "�������� ����������",
                    CodeName = "SOURCE_DESC",
                    ColumnIndex = 44,
                    SearchCritetias = new List<SearchCritetia>(new[]
                    {
                        "��������_����������", "��������"
                    }.Select(s => new SearchCritetia {Text = s}).ToList())
                },
                new TemplateColumn
                {
                    Name = "������ �� �������� ����������",
                    CodeName = "URL_SALE",
                    ColumnIndex = 45,
                    SearchCritetias = new List<SearchCritetia>(new[]
                    {
                        "������_��_��������_����������", "������"
                    }.Select(s => new SearchCritetia {Text = s}).ToList())
                },
                new TemplateColumn {Name = "������������ ��������", CodeName = "SELLER", ColumnIndex = 46},
                new TemplateColumn {Name = "��������������-�������� �����", CodeName = "OKOPF", ColumnIndex = 47},
                new TemplateColumn {Name = "����� ����� � ���� ��������", CodeName = "URL_INFO", ColumnIndex = 48},
                new TemplateColumn
                {
                    Name = "������� ��������",
                    CodeName = "CONTACTS",
                    ColumnIndex = 49,
                    SearchCritetias = new List<SearchCritetia>(new[]
                    {
                        "�������_��������", "��������", "�������", "��������"
                    }.Select(s => new SearchCritetia {Text = s}).ToList())
                },
                new TemplateColumn
                {
                    Name = "���� ���������� ����������",
                    CodeName = "DATE_RESEARCH",
                    ColumnIndex = 50,
                    SearchCritetias = new List<SearchCritetia>(new[]
                    {
                        "����_����������_����������", "����_����������", "����"
                    }.Select(s => new SearchCritetia {Text = s}).ToList())
                },
                new TemplateColumn {Name = "���� ������ �� �����", CodeName = "DATE_IN_BASE", ColumnIndex = 51},
                new TemplateColumn {Name = "������������", CodeName = "ACTUAL", ColumnIndex = 52},
                new TemplateColumn
                {
                    Name = "���� ��������",
                    CodeName = "DATE_IS_RINGING",
                    ColumnIndex = 53,
                    SearchCritetias = new List<SearchCritetia>(new[]
                    {
                        "��������"
                    }.Select(s => new SearchCritetia {Text = s}).ToList())
                },
                new TemplateColumn {Name = "��������� ��������", CodeName = "RESULT", ColumnIndex = 54},
                new TemplateColumn
                {
                    Name = "���������� (�����������) ��������������",
                    CodeName = "ADDITIONAL",
                    ColumnIndex = 55
                },
                new TemplateColumn {Name = "�����������", CodeName = "COMMENT", ColumnIndex = 56},
                new TemplateColumn {Name = "������������ � �����������", CodeName = "ASSOCIATIONS", ColumnIndex = 57},
                new TemplateColumn
                {
                    Name = "���� ��������",
                    CodeName = "DATE_PARSING",
                    ColumnIndex = 58,
                    SearchCritetias = new List<SearchCritetia>(new[]
                    {
                        "�������"
                    }.Select(s => new SearchCritetia {Text = s}).ToList())
                },
                new TemplateColumn {Name = "���������", CodeName = "LAND_MARK", ColumnIndex = 59},
                new TemplateColumn {Name = "������������", CodeName = "SNT", ColumnIndex = 60}
            };

            var landWb = new TemplateWorkbook {WorkbookType = XlTemplateWorkbookType.LandProperty};
            landWb.Columns.AddRange(columns);

            context.TemplateWorkbooks.Add(landWb);

            context.SaveChanges();
        }
    }
}