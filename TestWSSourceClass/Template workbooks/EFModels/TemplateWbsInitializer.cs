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
            InitializeComerceWorkbook(context);
        }

        private void InitializeComerceWorkbook(TemplateWbsContext context)
        {
            #region columns
            var columns = new[]
            {
                new TemplateColumn {CodeName = "ID", Name = "����������_�����", ColumnIndex = 1},
                new TemplateColumn
                {
                    CodeName = "SUBJECT",
                    Name = "�������_����������_���������",
                    ColumnIndex = 2,
                    SearchCritetias = new List<SearchCritetia>(new[]
                    {
                        "�������_����������_���������", "�������", "���������", "�������", "����","������"
                    }.Select(s => new SearchCritetia {Text = s}).ToList())
                },
                new TemplateColumn {CodeName = "REGION", Name = "�������������_�����������_(�����)", ColumnIndex = 3,
                    SearchCritetias = new List<SearchCritetia>(new[]
                    {
                        "��������������", "�����", "�����", "������"
                    }.Select(s => new SearchCritetia {Text = s}).ToList())},
                new TemplateColumn {CodeName = "SETTLEMENT", Name = "���������", ColumnIndex = 4,
                    SearchCritetias = new List<SearchCritetia>(new[]
                    {
                        "��������", "�����","�����"
                    }.Select(s => new SearchCritetia {Text = s}).ToList())},
                new TemplateColumn {CodeName = "CITY", Name = "����������_�����", ColumnIndex = 5},
                new TemplateColumn {CodeName = "CITY_TYPE", Name = "���_�����������_������", ColumnIndex = 6},
                new TemplateColumn {CodeName = "VGT", Name = "���������������_����������", ColumnIndex = 7,
                    SearchCritetias = new List<SearchCritetia>(new[]
                    {
                        "���������","��������"
                    }.Select(s => new SearchCritetia {Text = s}).ToList())},
                new TemplateColumn {CodeName = "STREET", Name = "�����", ColumnIndex = 8,
                    SearchCritetias = new List<SearchCritetia>(new[]
                    {
                        "����"
                    }.Select(s => new SearchCritetia {Text = s}).ToList())},
                new TemplateColumn {CodeName = "STREET_TYPE", Name = "���_�����", ColumnIndex = 9},
                new TemplateColumn {CodeName = "HOUSE_NUM", Name = "���", ColumnIndex = 10},
                new TemplateColumn {CodeName = "LETTER", Name = "������", ColumnIndex = 11},
                new TemplateColumn {CodeName = "BUILDING", Name = "������", ColumnIndex = 12},
                new TemplateColumn {CodeName = "STRUCTURE", Name = "��������", ColumnIndex = 13},
                new TemplateColumn {CodeName = "ESTATE", Name = "��������", ColumnIndex = 14},
                new TemplateColumn {CodeName = "LONGITUDE", Name = "�������", ColumnIndex = 15},
                new TemplateColumn {CodeName = "LATITUDE", Name = "������", ColumnIndex = 16},
                new TemplateColumn
                {
                    CodeName = "DIST_REG_CENTER",
                    Name = "�����������_��_�������������_������",
                    ColumnIndex = 17,
                    SearchCritetias = new List<SearchCritetia>(new[]
                    {
                        "�����������", "�����"
                    }.Select(s => new SearchCritetia {Text = s}).ToList())
                },
                new TemplateColumn
                {
                    CodeName = "CADASTRE_NUM",
                    Name = "�����������_�����_����������_�������",
                    ColumnIndex = 18
                },
                new TemplateColumn {CodeName = "METRO", Name = "�������_�����", ColumnIndex = 19},
                new TemplateColumn {CodeName = "METRO_DISTMIN", Name = "��_�����_�����", ColumnIndex = 20},
                new TemplateColumn {CodeName = "TRANSPORT", Name = "������_�����������", ColumnIndex = 21},
                new TemplateColumn {CodeName = "SEGMENT", Name = "�������", ColumnIndex = 22},
                new TemplateColumn {CodeName = "BUILDING_TYPE", Name = "���_���������", ColumnIndex = 23},
                new TemplateColumn {CodeName = "CENTER_CodeName", Name = "������������_������", ColumnIndex = 24},
                new TemplateColumn {CodeName = "OBJECT_TYPE", Name = "���_�������", ColumnIndex = 25},
                new TemplateColumn {CodeName = "OBJECT_PURPOSE", Name = "����������_�������", ColumnIndex = 26},
                new TemplateColumn {CodeName = "CLASS_TYPE", Name = "���������������_�����", ColumnIndex = 27},
                new TemplateColumn {CodeName = "OPERATION", Name = "��������", ColumnIndex = 28,
                    SearchCritetias = new List<SearchCritetia>(new[]
                    {
                        "���_������","��� ������","�����"
                    }.Select(s => new SearchCritetia {Text = s}).ToList())},
                new TemplateColumn {CodeName = "SALE_PRICE", Name = "���� _�������", ColumnIndex = 29,
                    SearchCritetias = new List<SearchCritetia>(new[]
                    {
                        "���������", "�����", "����", "������"
                    }.Select(s => new SearchCritetia {Text = s}).ToList())},
                new TemplateColumn {CodeName = "RENT_RATE", Name = "��������_�����", ColumnIndex = 30},
                new TemplateColumn {CodeName = "AREA", Name = "�������", ColumnIndex = 31,
                    SearchCritetias = new List<SearchCritetia>(new[]
                    {
                        "�������", "������", "����"
                    }.Select(s => new SearchCritetia {Text = s}).ToList())},
                new TemplateColumn {CodeName = "PRICE_FOR_UNIT", Name = "����_��_�2", ColumnIndex = 32},
                new TemplateColumn {CodeName = "OPERATING_COSTS", Name = "����������������_�������", ColumnIndex = 33},
                new TemplateColumn {CodeName = "FLOOR", Name = "����", ColumnIndex = 34},
                new TemplateColumn {CodeName = "FLOOR_QNT_MIN", Name = "���������_�����������", ColumnIndex = 35},
                new TemplateColumn {CodeName = "FLOOR_QNT_MAX", Name = "���������_������������", ColumnIndex = 36},
                new TemplateColumn {CodeName = "YEAR_BUILD", Name = "���_���������", ColumnIndex = 37},
                new TemplateColumn {CodeName = "MATERIAL_WALL", Name = "��������_����", ColumnIndex = 38},
                new TemplateColumn {CodeName = "HEIGHT_FLOOR", Name = "������_�������", ColumnIndex = 39},
                new TemplateColumn {CodeName = "COLUMN_DIST", Name = "���_������", ColumnIndex = 40},
                new TemplateColumn {CodeName = "LAYOUT", Name = "����������", ColumnIndex = 41},
                new TemplateColumn {CodeName = "ROOM_QNT", Name = "����������_������", ColumnIndex = 42},
                new TemplateColumn {CodeName = "AREA_TOTAL", Name = "�������_�����", ColumnIndex = 43},
                new TemplateColumn
                {
                    CodeName = "AREA_LOT",
                    Name = "�������_����������_�������_�������",
                    ColumnIndex = 44
                },
                new TemplateColumn {CodeName = "CONDITION", Name = "���������", ColumnIndex = 45},
                new TemplateColumn {CodeName = "SECURITY", Name = "������������", ColumnIndex = 46},
                new TemplateColumn {CodeName = "FLOOR_LOAD", Name = "���������� �������� �� ���", ColumnIndex = 47},
                new TemplateColumn {CodeName = "CONDITIONING", Name = "�����������������", ColumnIndex = 48},
                new TemplateColumn {CodeName = "VENT", Name = "����������", ColumnIndex = 49},
                new TemplateColumn
                {
                    Name = "�������������",
                    CodeName = "SYSTEM_GAS",
                    ColumnIndex = 50,
                    SearchCritetias = new List<SearchCritetia>(new[]
                    {
                        "�������������", "���������", "���","����������"
                    }.Select(s => new SearchCritetia {Text = s}).ToList())
                },
                new TemplateColumn
                {
                    Name = "�������������",
                    CodeName = "SYSTEM_WATER",
                    ColumnIndex = 51,
                    SearchCritetias = new List<SearchCritetia>(new[]
                    {
                        "�������������", "��������", "�����", "���"
                    }.Select(s => new SearchCritetia {Text = s}).ToList())
                },
                new TemplateColumn
                {
                    Name = "�����������",
                    CodeName = "SYSTEM_SEWERAGE",
                    ColumnIndex = 52,
                    SearchCritetias = new List<SearchCritetia>(new[]
                    {
                        "�����������", "���������", "�������", "�����"
                    }.Select(s => new SearchCritetia {Text = s}).ToList())
                },
                new TemplateColumn
                {
                    Name = "����������������",
                    CodeName = "SYSTEM_ELECTRICITY",
                    ColumnIndex = 53,
                    SearchCritetias = new List<SearchCritetia>(new[]
                    {
                        "����������������", "�����������","����������", "��������", "�������", "���"
                    }.Select(s => new SearchCritetia {Text = s}).ToList())
                },
                new TemplateColumn
                {
                    Name = "��������������",
                    CodeName = "HEAT_SUPPLY",
                    ColumnIndex = 54,
                    SearchCritetias = new List<SearchCritetia>(new[]
                    {
                        "��������������", "���������", "����", "�����", "�����"
                    }.Select(s => new SearchCritetia {Text = s}).ToList())
                },
                new TemplateColumn {CodeName = "TRAIN", Name = "�/�_�����", ColumnIndex = 55},
                new TemplateColumn {CodeName = "ROAD", Name = "������", ColumnIndex = 56},
                new TemplateColumn {CodeName = "DESCRIPTION", Name = "��������", ColumnIndex = 57,
                    SearchCritetias = new List<SearchCritetia>(new[]
                    {
                        "��������"
                    }.Select(s => new SearchCritetia {Text = s}).ToList())},
                new TemplateColumn {CodeName = "SOURCE_DESC", Name = "��������_����������",
                    SearchCritetias = new List<SearchCritetia>(new[]
                    {
                        "��������_����������", "��������","������"
                    }.Select(s => new SearchCritetia {Text = s}).ToList()), ColumnIndex = 58},
                new TemplateColumn {CodeName = "SOURCE_LINK", Name = "������_��_��������_����������", ColumnIndex = 59,
                    SearchCritetias = new List<SearchCritetia>(new[]
                    {
                        "������_��_��������_����������", "������","URL"
                    }.Select(s => new SearchCritetia {Text = s}).ToList())},
                new TemplateColumn {CodeName = "CONTACTS", Name = "��������", ColumnIndex = 60,
                    SearchCritetias = new List<SearchCritetia>(new[]
                    {
                        "�������_��������", "��������", "�������", "��������"
                    }.Select(s => new SearchCritetia {Text = s}).ToList())},
                new TemplateColumn {CodeName = "DATE_RESEARCH", Name = "����_�����_����������", ColumnIndex = 61,
                    SearchCritetias = new List<SearchCritetia>(new[]
                    {
                        "����_����������_����������", "����_����������", "����"
                    }.Select(s => new SearchCritetia {Text = s}).ToList())},
                new TemplateColumn
                {
                    Name = "���� ��������",
                    CodeName = "DATE_PARSING",
                    ColumnIndex = 62,
                    SearchCritetias = new List<SearchCritetia>(new[]
                    {
                        "�������"
                    }.Select(s => new SearchCritetia {Text = s}).ToList())
                },

            };

#endregion

            var commerceWb = new TemplateWorkbook { WorkbookType = XlTemplateWorkbookType.CommerceProperty };
            commerceWb.Columns.AddRange(columns);

            context.TemplateWorkbooks.Add(commerceWb);

            context.SaveChanges();
        }

        private void InitializeLandWorkbook(TemplateWbsContext context)
        {
            #region Columns
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
                        "�������_����������_���������", "�������", "���������", "�������", "����","������"
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
                        "��������", "�����","�����"
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
                new TemplateColumn {Name = "��������� �����", CodeName = "VGT", ColumnIndex = 8,
                    SearchCritetias = new List<SearchCritetia>(new[]
                    {
                        "���������","��������"
                    }.Select(s => new SearchCritetia {Text = s}).ToList())},
                new TemplateColumn {Name = "������������ ��������� �������", CodeName = "STREET", ColumnIndex = 9,
                    SearchCritetias = new List<SearchCritetia>(new[]
                    {
                        "����"
                    }.Select(s => new SearchCritetia {Text = s}).ToList())},
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
                new TemplateColumn {Name = "��������", CodeName = "OPERATION", ColumnIndex = 23,
                    SearchCritetias = new List<SearchCritetia>(new[]
                    {
                        "��� ������"
                    }.Select(s => new SearchCritetia {Text = s}).ToList())},
                new TemplateColumn
                {
                    Name = "����� �� �������",
                    CodeName = "LAW_NOW",
                    ColumnIndex = 24,
                    SearchCritetias = new List<SearchCritetia>(new[]
                    {
                        "���_�����", "��� �����", "�����", "����","��� ��������"
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
                        "���������", "�����", "����", "������","�����","����� �����"
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
                        "�������������", "���������", "���","����������"
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
                        "����������������", "�����������","����������" ,"��������", "�������", "����"
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
                new TemplateColumn {Name = "������� �������� �� �������", CodeName = "OBJECT", ColumnIndex = 38,
                    SearchCritetias = new List<SearchCritetia>(new[]
                    {
                        "��������","���������"
                    }.Select(s => new SearchCritetia {Text = s}).ToList())},
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
                        "��������_����������", "��������","������"
                    }.Select(s => new SearchCritetia {Text = s}).ToList())
                },
                new TemplateColumn
                {
                    Name = "������ �� �������� ����������",
                    CodeName = "URL_SALE",
                    ColumnIndex = 45,
                    SearchCritetias = new List<SearchCritetia>(new[]
                    {
                        "������_��_��������_����������", "������","URL"
                    }.Select(s => new SearchCritetia {Text = s}).ToList())
                },
                new TemplateColumn {Name = "������������ ��������", CodeName = "SELLER", ColumnIndex = 46,
                    SearchCritetias = new List<SearchCritetia>(new[]
                    {
                        "�������_��������", "��������", "��������","��������"
                    }.Select(s => new SearchCritetia {Text = s}).ToList())},
                new TemplateColumn {Name = "��������������-�������� �����", CodeName = "OKOPF", ColumnIndex = 47},
                new TemplateColumn {Name = "����� ����� � ���� ��������", CodeName = "URL_INFO", ColumnIndex = 48},
                new TemplateColumn
                {
                    Name = "��������",
                    CodeName = "CONTACTS",
                    ColumnIndex = 49,
                    SearchCritetias = new List<SearchCritetia>(new[]
                    {
                        "�������_��������", "��������", "���","����"
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
                new TemplateColumn {Name = "��������", CodeName = "LAND_MARK", ColumnIndex = 59},
                new TemplateColumn {Name = "������������", CodeName = "SNT", ColumnIndex = 60}
            };
#endregion

            var landWb = new TemplateWorkbook {WorkbookType = XlTemplateWorkbookType.LandProperty};
            landWb.Columns.AddRange(columns);

            context.TemplateWorkbooks.Add(landWb);

            context.SaveChanges();
        }
    }
}