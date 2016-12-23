using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace NPOIHelper.NPOI.Abstract
{
    public abstract class TableFooter
    {
        public TableFooter()
        {
            Rows = new List<Row>();
        }
        public IList<Row> Rows { get; set; }

        public string DefaultTableFooterCellStyle { get; set; }

        public bool IsNull()
        {
            return Rows.Count() == 0 ? true : false;
        }

        abstract public void AddRow(Row row);
    }
}
