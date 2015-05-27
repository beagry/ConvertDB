using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Converter.Template_workbooks.EFModels
{
    class TemplateWbsProxy
    {
        private TemplateWbsContext db;

        public TemplateWbsProxy()
        {
            db = new TemplateWbsContext();
        }

        public IEnumerable<TemplateWorkbook> TemplateWorkbooks { get; set; }
    }
}
