using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace NPOIHelper.NPOI
{
    public class Row
    {
        public Row()
        {
            Columns = new List<Cell>();
        }
        public Row(IList<Cell> cols)
        {
            Columns = cols;
        }
        public bool IsHead { get; set; }
        public IList<Cell> Columns { get; set; }
        
    }
}
