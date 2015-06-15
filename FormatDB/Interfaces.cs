using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Converter.Template_workbooks;

namespace Formater
{
    public interface IFormatDbParams
    {
        string Path { get; set; }
        XlTemplateWorkbookType WorkbookType { get; set; }
        bool DoDescription { get; set; }

        ISupportWorkbook CatalogSupportWorkbook { get; set; }
        ISupportWorkbook OktmoSupportWorkbook { get; set; }
        ISupportWorkbook VgtCatalogSupportWorkbook { get; set; }
        ISupportWorkbook SubjectSourceSupportWorkbook { get; set; }
    }

    public interface ISupportWorkbook
    {
        string Path { get; set; }

        string SelectedWorksheet { get; set; }
    }
}
