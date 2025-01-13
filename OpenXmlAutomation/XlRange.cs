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
        public List<List<XlCell>> Cells { get; private set; }

        public int Width => Cells.Count;

        public int Height => Cells.Count > 0 ? Cells[0].Count : 0;

        internal XlRange(XlSheet s, string tlbr, List<List<XlCell>> cells) 
        { 
            sheet = s;
            cellRange = tlbr; 
            Cells = cells;
        }
        
    }
}
