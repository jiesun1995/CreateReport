using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ConsoleApp1.Attributes.ReportAttribute
{
    public class CellAttribute:Attribute
    {
        public string Name { get; set; }
        public string ColSpan { get; set; }
        public string RowSpan { get; set; }
        public string GroupName { set; get; }
    }
}
