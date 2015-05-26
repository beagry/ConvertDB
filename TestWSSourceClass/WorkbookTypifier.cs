using System.Collections.Generic;
using System.Linq;
using Converter.Template_workbooks;
using ExcelRLibrary;
using Microsoft.Office.Interop.Excel;

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
        public Workbook CombineToSingleWorkbook()
        {
            var helper = new ExcelHelper();

            //������� ������ �����
            var newWb = helper.CreateNewWorkbook();
            var ws = newWb.Worksheets[1] as Worksheet;
            
            //�������� �����
            var templateHead = new T().TemplateColumns.ToDictionary(k => k.Index, v => v.CodeName);
            ws.WriteHead(templateHead);

            var wsWriter = new WorksheetFiller(ws, RulesDictionary);
            
            //���������� ���������� ����� �� ������ ����� �� ������
            foreach (var openWs in helper.GetWorkbooks(WorkbooksPaths).Select(wb => wb.Worksheets[1]).Cast<Worksheet>())
                wsWriter.InsertWorksheet(openWs);

            return newWb;
        }        
    }
}