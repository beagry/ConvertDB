using System.Collections.Generic;

namespace Converter.Template_workbooks
{
    public class JustColumn
    {

        public int Index { get; set; }

        public string Description { get; private set; }

        public string CodeName { get; set; }

        public List<string> Examples { get; set; }




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
    }
}