using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace NPOIHelper.NPOI.Abstract
{
    public abstract class TableBody
    {
        public TableBody()
        {
            Rows = new List<Row>();
        }
        public IList<Row> Rows { get; set; }

        private string defaultTableBodyCellStyle;
        public string DefaultTableBodyCellStyle
        {
            get { return defaultTableBodyCellStyle; }
            set { defaultTableBodyCellStyle = value; }
        }

        abstract public void AddRow(Row row);
    }
}
