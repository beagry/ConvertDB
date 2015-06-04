using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;

namespace TestWSSourceClass
{
    static class TemplateWorkbook
    {
        private static readonly List<JustColumn> columns = new List<JustColumn>
            #region Columns Initialize
        {
            new JustColumn("SUBJECT", "������� ���������� ���������",2),
            new JustColumn("REGION", "������������� ����������� (�����)",3),
            new JustColumn("SETTLEMENT", "���������",4),
            new JustColumn("NEAR_CITY", "��������� ���������� �����",5),
            new JustColumn("TERRITORY_TYPE", "��� ���������� ����������� ������",6),
            new JustColumn("IN_CITY", "������ ���������� � �������� ����������� ������",7),
            new JustColumn("VGT", "��������� �����",8),
            new JustColumn("STREET", "������������ ��������� �������",9),
            new JustColumn("STREET_TYPE", "��� ��������� �������",10),
            new JustColumn("HOUSE_NUM", "���",11),
            new JustColumn("LETTER", "������",12),
            new JustColumn("BUILDING", "������",13),
            new JustColumn("STRUCTURE", "��������",14),
            new JustColumn("ESTATE", "��������",15),
            new JustColumn("LONGITUDE", "�������",16),
            new JustColumn("LATITUDE", "������",17),
            new JustColumn("HIGHWAY", "������",18),
            new JustColumn("DIST_REG_CENTER", "���������� �� ������������� ������",19),
            new JustColumn("DIST_NEAR_CITY", "���������� �� ���������� ����������� ������",20),
            new JustColumn("CADASTRE_NUM", "����������� ����� ���������� �������",21),
            new JustColumn("OFFER_DEAL", "����������� (������)",22),
            new JustColumn("OPERATION", "��������",23),
            new JustColumn("LAW_NOW", "����� �� �������",24),
            new JustColumn("SALE_TYPE", "������ ����������",25),
            new JustColumn("RENTAL_PERIOD", "���� ������",26),
            new JustColumn("PRICE", "���� ����������� (������)",27),
            new JustColumn("RENT_RATE", "�������� �����",28),
            new JustColumn("AREA_LOT", "�������",29),
            new JustColumn("LAND_CATEGORY", "��������� ������",30),
            new JustColumn("PERMITTED_USE", "��� ������������ �������������",31),
            new JustColumn("PERMITTED_USE_TEXT", "��� ������������ ������������� �����",32),
            new JustColumn("SYSTEM_GAS", "�������������",33),
            new JustColumn("SYSTEM_WATER", "�������������",34),
            new JustColumn("SYSTEM_SEWERAGE", "�����������",35),
            new JustColumn("SYSTEM_ELECTRICITY", "����������������",36),
            new JustColumn("HEAT_SUPPLY", "��������������",37),
            new JustColumn("OBJECT", "������� �������� �� �������",38),
            new JustColumn("SURFACE", "�������� ��������",39),
            new JustColumn("ROAD", "������",40),
            new JustColumn("RELIEF", "������",41),
            new JustColumn("VEGETATION", "������������ ������",42),
            new JustColumn("DESCRIPTION", "��������",43),
            new JustColumn("SOURCE_DESC", "�������� ����������",44),
            new JustColumn("URL_SALE", "������ �� �������� ����������",45),
            new JustColumn("SELLER", "������������ ��������",46),
            new JustColumn("OKOPF", "��������������-�������� �����",47),
            new JustColumn("URL_INFO", "����� ����� � ���� ��������",48),
            new JustColumn("CONTACTS", "��������",49),
            new JustColumn("DATE_RESEARCH", "���� ���������� ����������",50),
            new JustColumn("DATE_IN_BASE", "���� ������ �� �����",51),
            new JustColumn("ACTUAL", "������������",52),
            new JustColumn("DATE_IS_RINGING", "���� ��������",53),
            new JustColumn("RESULT", "��������� ��������",54),
            new JustColumn("ADDITIONAL", "���������� (�����������) ��������������",55),
            new JustColumn("COMMENT", "�����������",55),


        };

            #endregion
        public static IEnumerable<JustColumn> TemplateColumns
        {
            get { return columns; }
        }

        public static int GetColumnByCode(string name)
        {
            int column = 0;
            JustColumn firstOrDefault = columns.FirstOrDefault(x => x.Code == name);
            if (firstOrDefault != null)
                column = firstOrDefault.Index;
            return column;
        }

        public static Dictionary<string, int> GroupWorkBooksByHead(IEnumerable<string> workbooksPaths)
        {
            var xlApplication = GetExcelApplication();
            var resultDictionary = new Dictionary<string, int>();


            var wsTypes = new List<WSType>();
            var n = 1;
            foreach (var s in workbooksPaths)
            {
                Process.Start(s);
                var workbook = Enumerable.Cast<Microsoft.Office.Interop.Excel.Workbook>(xlApplication.Workbooks)
                    .First(x => x.Name == System.IO.Path.GetFileName(s));

#if DEBUG
                Debug.Assert(workbook != null);
#endif

                Microsoft.Office.Interop.Excel.Worksheet worksheet = workbook.Worksheets[1];
                var head = new List<string>();

                var lastUsedColumn = worksheet.Cells.SpecialCells(Microsoft.Office.Interop.Excel.XlCellType.xlCellTypeLastCell).Column;
                var headRow = worksheet.Range[worksheet.Cells[1, 1], worksheet.Cells[1, lastUsedColumn]];

                foreach (Microsoft.Office.Interop.Excel.Range cell in headRow)
                    if (!String.IsNullOrEmpty(cell.Value2))
                        head.Add(cell.Value2.ToString());

                if (wsTypes.Any(x => x.Heads.SequenceEqual(head)))
                {
                    resultDictionary.Add(s, wsTypes.First(x => x.Heads.SequenceEqual(head)).GroupNumber);
                }
                else
                {
                    wsTypes.Add(new WSType { Heads = head, GroupNumber = n });
                    resultDictionary.Add(s, n);
                    n++;
                }
                workbook.Close(false);
            }
            return resultDictionary;
        }

        public static Microsoft.Office.Interop.Excel.Application GetExcelApplication()
        {
            Microsoft.Office.Interop.Excel.Application xlApplication = null;
            try
            {
                xlApplication = (Microsoft.Office.Interop.Excel.Application)Marshal.GetActiveObject("Excel.Application");
            }
            catch (COMException exception)
            {
                if (xlApplication == null)
                {
                    xlApplication = new Microsoft.Office.Interop.Excel.Application(){ Visible = false };
                }
                else
                {
                    throw exception;
                }
            }
            return xlApplication;
        }
    }
}