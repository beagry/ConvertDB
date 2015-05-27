using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations.Schema;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Converter.Template_workbooks.EFModels
{
    public class TemplateWorkbook
    {
        public TemplateWorkbook()
        {
            Columns = new List<TemplateColumn>();
        }
        public int Id { get; set; }

        [NotMapped]
        public string Name { get { return "SuperName"; } }

        public XlTemplateWorkbookType WorkbookType { get; set; }

        public virtual List<TemplateColumn> Columns { get; set; } 
    }

    public class TemplateColumn
    {
        public TemplateColumn()
        {
            SearchCritetias = new List<SearchCritetia>();
            BindedColumns = new List<BindedColumn>();
        }
        public int Id { get; set; }
        public string Name { get; set; }
        public string CodeName { get; set; }
        public int ColumnIndex { get; set; }


        public virtual List<SearchCritetia> SearchCritetias { get; set; }

        public virtual List<BindedColumn> BindedColumns { get; set; }
    }

    public class BindedColumn
    {
        public int Id { get; set; }
        public string Name { get; set; }
    }

    public class SearchCritetia
    {
        public int Id { get; set; }
        public string Text { get; set; }
    }
}
