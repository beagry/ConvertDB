using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
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

        ISupportWorkbookViewModel CatalogSupportWorkbook { get; set; }
        ISupportWorkbookViewModel OktmoSupportWorkbook { get; set; }
        ISupportWorkbookViewModel VgtCatalogSupportWorkbook { get; set; }
        ISupportWorkbookViewModel SubjectSourceSupportWorkbook { get; set; }
        ISupportWorkbookViewModel KladrWorkbook { get; set; }
    }

    public interface ISupportWorkbook
    {
        string Path { get; set; }

        string SelectedWorksheet { get; set; }
    }

    public interface ISupportWorkbookViewModel : ISupportWorkbook, INotifyPropertyChanged
    {
        ObservableCollection<string> Worksheets { get; set; }

        bool HasWorksheets { get; set; }
        bool TaskInProgress { get; set; }
    }
}
