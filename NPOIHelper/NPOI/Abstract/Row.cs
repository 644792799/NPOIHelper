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
            HaveRowBreak = false;
        }
        /// <summary>
        /// 是否是横向表头
        /// </summary>
        public bool IsRowHead { get; set; }
        /// <summary>
        /// 是否是竖向表头
        /// </summary>
        public bool IsColHead { get; set; }
        /// <summary>
        /// 是否包含分页符
        /// </summary>
        public bool HaveRowBreak { get; set; }
        public IList<Cell> Cells { get; set; }
        private short height = 25;
        public short Height 
        {
            get { return height; }
            set { height = value; } 
        }
        /// <summary>
        /// 创建单元格
        /// </summary>
        /// <returns></returns>
        abstract public Cell CreateCell();
        abstract public void AddCell(Cell cell);
        abstract public void AddCell(int columnIndex);
        abstract public void DeleteCell(int columnIndex);
        abstract public void DeleteCells(int fromColumnIndex, int toColumnIndex);
    }
}
