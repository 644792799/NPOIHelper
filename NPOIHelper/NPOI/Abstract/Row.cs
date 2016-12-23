using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace NPOIHelper.NPOI.Abstract
{
    public abstract class Row
    {
        const short DEFAULT_ROW_HEIGHT = 25;
        public Row()
        {
            Cells = new List<Cell>();
            HaveRowBreak = false;
        }
        /// <summary>
        /// 是否包含分页符
        /// </summary>
        public bool HaveRowBreak { get; set; }
        public IList<Cell> Cells { get; set; }
        private short height = DEFAULT_ROW_HEIGHT;
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
    }
}
