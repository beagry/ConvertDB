using System;
using System.Collections.Generic;
using System.Linq;

namespace Converter.Template_workbooks
{
    public class TemplateWorkbook
    {
        protected List<JustColumn> columns;
        public String UnUsedColumnCode
    {
            get { return "UNUS"; }
    }

        public IEnumerable<JustColumn> TemplateColumns
        {
            get { return columns; }
        }

        public int GetColumnByCode(string name)
        {
            int column = 0;
            JustColumn firstOrDefault = columns.FirstOrDefault(x => x.CodeName == name);
            if (firstOrDefault != null)
                column = firstOrDefault.Index;
            return column;
        }

        

        public static Microsoft.Office.Interop.Excel.Workbook GetTemplateWorkbook()
        {
            throw new Exception("Метод не готов!");
            //Excel.Workbook workbook = (new Excel.Application()).Workbooks.Add();
            //return workbook;
        }
    }
    public class JustColumn
    {
        public JustColumn(string codename, string description, int index)
        {
            Index = index;
            Description = description;
            CodeName = codename;
        }

        public JustColumn(string description, int index)
        {
            Index = index;
            Description = description;
        }
        public int Index { get; set; }

        public string Description { get; private set; }

        public string CodeName { get; set; }

        public List<string> Examples { get; set; }
    }
    public class WSType
    {
        public List<string> Heads { get; set; }
        public int GroupNumber { get; set; }

    }
}