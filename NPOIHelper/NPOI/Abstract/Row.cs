using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace NPOIHelper.NPOI.Abstract
{
    public abstract class Row
    {
        public Row()
        {
            Cells = new List<Cell>();
            IsRowHead = false;
            IsColHead = false;
        }
        public bool IsRowHead { get; set; }
        public bool IsColHead { get; set; }
        public IList<Cell> Cells { get; set; }
        private short height = 25;
        public short Height 
        {
            get { return height; }
            set { height = value; } 
        }

        abstract public Cell CreateCell();
        abstract public void AddCell(Cell cell);
        abstract public void AddCell(int columnIndex);
        abstract public void DeleteCell(int columnIndex);
        abstract public void DeleteCells(int fromColumnIndex, int toColumnIndex);
    }
}
