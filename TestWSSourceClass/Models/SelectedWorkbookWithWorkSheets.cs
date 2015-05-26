using ExcelRLibrary;

namespace Converter.Models
{
    public class SelectedWorkbookWithWorkSheets : SelectedWorkbook
    {

        public WorksheetInfo WorksheetInfo { get; private set; }

        public SelectedWorkbookWithWorkSheets():base()
        {
            
        }

        public SelectedWorkbookWithWorkSheets(SelectedWorkbook simpleWorkbook, WorksheetInfo wsInfo)
        {
            Name = simpleWorkbook.Name;
            Path = simpleWorkbook.Path;
            WorksheetInfo = wsInfo;
        }

    }
}