using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;

namespace Formater
{
    public class JustColumn
    {
        public JustColumn(string name, string description, int index):this(description,index)
        {
            Code = name;
        }

        public JustColumn(string description, int index)
        {
            Index = index;
            Name = description;
        }
        public int Index { get; set; }

        public string Name { get; private set; }

        public string Code { get; set; }

        public List<string> Examples { get; set; }
    }
    public class WSType
    {
        public List<string> Heads { get; set; }
        public int GroupNumber { get; set; }

    }
}
