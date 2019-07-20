using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelCharts2Powerpoint
{
    class PresentationItem
    {
        public string OutputFileName="";
        public List<Attribute> Attributes = new List<Attribute>();
    }
    class Attribute
    {
        public string type = "";
        public string name = "";
        public string value = "";
        public Attribute(string type, string name, string value)
        {
            this.type = type;
            this.name = name;
            this.value = value;
        }
    }
}
