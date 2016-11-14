using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace NPOIHelper.NPOI.Abstract
{
    public abstract class Header
    {
        public Header()
        {
            Rows = new List<Row>();
        }
        public IList<Row> Rows { get; set; }
    }
}
