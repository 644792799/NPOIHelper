using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace NPOIHelper.NPOI.Abstract
{
    public abstract class TableHeader
    {
        public TableHeader()
        {
            Rows = new List<Row>();
        }
        public IList<Row> Rows { get; set; }

        private string defaultTableHeaderCellStyle;
        public string DefaultTableHeaderCellStyle
        {
            get { return defaultTableHeaderCellStyle; }
            set { defaultTableHeaderCellStyle = value; }
        }

        abstract public void AddRow(Row row);
    }
}
