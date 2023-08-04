using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ConsoleApp1.Attributes.ReportAttribute
{
    public class ColNameAttribute:Attribute
    {
        public string Name { get; set; }
    }
}
