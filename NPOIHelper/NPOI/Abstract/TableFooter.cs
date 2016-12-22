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

        private string defaultTableFooterCellStyle;
        public string DefaultTableFooterCellStyle
        {
            get { return defaultTableFooterCellStyle; }
            set { defaultTableFooterCellStyle = value; }
        }

        abstract public void AddRow(Row row);
    }
}
