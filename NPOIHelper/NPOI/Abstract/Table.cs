using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace NPOIHelper.NPOI.Abstract
{
    public abstract class Table
    {
        public Header Header { get; set; }
        public string Title { get; set; }
        private short titleHeight = 30;
        public short TitleHeight {
            get { return titleHeight; }
            set { this.titleHeight = value; } 
        }
        public IList<Row> Rows { get; set; }
        public int ColumnCount { get; set; }
        public int[] ColumnWidths { get; set; }
        //public IList<string> Columns { get; set; }
        public Footer Footer { get; set; }

        public Table()
        {
            Rows = new List<Row>();
        }

        abstract public Row CreateRow(bool isHeader);
        abstract public void AddRow(Row row);
        abstract public void AddRow(int rowIndex);
        abstract public void DeleteRow(int rowIndex);
        abstract public void DeleteRows(int fromRowIndex, int toRowIndex);

    }
}
