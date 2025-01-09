using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OpenXmlAutomation
{
    /// <summary>
    /// Represents a range of cells
    /// </summary>
    public class XlRange
    {
        private string cellRange;
        private XlSheet sheet;
        public XlCell? TopLeft { get; set; }
        public XlCell? BottomRight { get; set; }
        internal XlRange(XlSheet s, string tlbr) 
        { 
            sheet = s;
            cellRange = tlbr; 
        }
    }
}
