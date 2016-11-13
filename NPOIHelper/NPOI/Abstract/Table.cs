using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace NPOIHelper.NPOI.Abstract
{
    public abstract class Table
    {
        public string Title { get; set; }
        public IList<Row> Rows { get; set; }
        public IList<string> Columns { get; set; }

        public Table()
        {
            Rows = new List<Row>();
        }

        abstract public Row CreateRow();
        abstract public void AddRow(Row row);
        abstract public void AddRow(int rowIndex);
        abstract public void DeleteRow(int rowIndex);
        abstract public void DeleteRows(int fromRowIndex, int toRowIndex);

    }
}
