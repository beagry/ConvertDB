using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Formater.SupportWorksheetsClasses
{
    class PlaceWithType
    {
        public PlaceWithType(string name, string type)
        {
            Name = name;
            Type = type;
        }

        public PlaceWithType()
        {
            
        }

        public string Name { get; set; }
        public string Type { get; set; }
    }
}
