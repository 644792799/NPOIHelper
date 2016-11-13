using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace NPOIHelper.NPOI.Abstract
{
    public abstract class Row
    {
        public bool IsHead { get; set; }
        public IList<Cell> Cells { get; set; }

        abstract public Cell CreateCell();
        abstract public void AddCell(Cell cell);
        abstract public void AddCell(int columnIndex);
        abstract public void DeleteCell(int columnIndex);
        abstract public void DeleteCells(int fromColumnIndex, int toColumnIndex);
    }
}
