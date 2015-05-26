using System.Collections.Generic;
using System.Linq;
using Converter.Template_workbooks;
using ExcelRLibrary;
using Microsoft.Office.Interop.Excel;
using OfficeOpenXml;
using DataTable = System.Data.DataTable;

namespace Converter
{
    /// <summary>
    /// ���� ������� ����� ����� �� ���������� ��������
    /// ����� ��� ����������� ���� � ������ �� ������ ���������� ������
    /// </summary>
    /// <typeparam name="T">�����-������. ������������ ��� ��������� Excel �������</typeparam>
    public class WorkbookTypifier<T> where T : TemplateWorkbook, new ()
    {
        public ICollection<string> WorkbooksPaths { get; set; }

        /// <summary>
        /// ������� ��� ����������� ����
        /// </summary>
        public Dictionary<string, List<string>> RulesDictionary { get; set; }



        public WorkbookTypifier(Dictionary<string, List<string>> rulesDictionary, ICollection<string> workbooksPaths)
        {
            RulesDictionary = rulesDictionary;
            this.WorkbooksPaths = workbooksPaths;
        }

        public WorkbookTypifier()
        {
            RulesDictionary = new Dictionary<string, List<string>>();
            WorkbooksPaths = new List<string>();
            
        }




        /// <summary>
        /// ����� ���������� ������ �����, ��������� �� ����������� ���� �� ���������� ��������
        /// </summary>
        /// <param name="workbooksPaths"></param>
        /// <returns></returns>
        public ExcelPackage CombineToSingleWorkbook()
        {
            var result = new ExcelPackage();
            var resultWS =  result.Workbook.Worksheets.Add("Combined");


            //����������� �������� ����
            var templateHead = new T().TemplateColumns.ToDictionary(k => k.Index, v => v.CodeName);
            resultWS.WriteHead(templateHead);

            var wsWriter = new WorksheetFiller(resultWS, RulesDictionary);


            var reader = new ExcelReader();
            foreach (
                var dt in
                    WorkbooksPaths.Select(p => reader.ReadExcelFile(p))
                        .Select(ds => ds.Tables.Cast<DataTable>().First()))
            {
                wsWriter.AppendDataTable(dt);
            }

            return result;
        }        
    }
}