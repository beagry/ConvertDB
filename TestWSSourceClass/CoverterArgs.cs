using System;
using System.Collections.Generic;

namespace Converter
{
    public class ConverterArgs:EventArgs
    {
        public ConverterArgs()
        {
            
        }
        public ConverterArgs(IEnumerable<string> workbooksPaths, XlTemplateWorkbookTypes workbooksType)
        {
            WorkbooksPaths = workbooksPaths;
            WorkbooksType = workbooksType;
        }

        public IEnumerable<string> WorkbooksPaths { get; set; }
        public XlTemplateWorkbookTypes WorkbooksType { get; set; }
    }
}
