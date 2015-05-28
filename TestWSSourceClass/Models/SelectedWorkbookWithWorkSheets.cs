using ExcelRLibrary;

namespace Converter.Models
{
    public class SelectedWorkbookWithWorkSheets : SelectedWorkbook
    {
        public SelectedWorkbookWithWorkSheets()
        {
        }

        public SelectedWorkbookWithWorkSheets(SelectedWorkbook simpleWorkbook, WorksheetInfo wsInfo)
        {
            Name = simpleWorkbook.Name;
            Path = simpleWorkbook.Path;
            WorksheetInfo = wsInfo;
        }

        public WorksheetInfo WorksheetInfo { get; private set; }
    }
}