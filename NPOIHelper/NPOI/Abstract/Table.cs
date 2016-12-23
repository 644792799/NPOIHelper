using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace NPOIHelper.NPOI.Abstract
{
    /// <summary>
    ///  ┏━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━┓
    ///  ┃                        TITLE                            ┃
    ///  ┣━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━┫
    ///  ┃                        HEADER                           ┃
    ///  ┣━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━┫
    ///  ┃    ┃    ┃    ┃    ┃TABLE HEADER┃    ┃    ┃    ┃    ┃    ┃
    ///  ┣━━━━╋━━━━╋━━━━╋━━━━╋━━━━━━━━━━━━╋━━━━╋━━━━╋━━━━╋━━━━╋━━━━┫
    ///  ┃    ┃    ┃    ┃    ┃TABLE   BODY┃    ┃    ┃    ┃    ┃    ┃
    ///  ┣━━━━╋━━━━╋━━━━╋━━━━╋━━━━━━━━━━━━╋━━━━╋━━━━╋━━━━╋━━━━╋━━━━┫
    ///  ┃    ┃    ┃    ┃    ┃TABLE   BODY┃    ┃    ┃    ┃    ┃    ┃
    ///  ┣━━━━╋━━━━╋━━━━╋━━━━╋━━━━━━━━━━━━╋━━━━╋━━━━╋━━━━╋━━━━╋━━━━┫
    ///  ┃    ┃    ┃    ┃    ┃TABLE FOOTER┃    ┃    ┃    ┃    ┃    ┃
    ///  ┣━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━┫
    ///  ┃                        FOOTER                           ┃
    ///  ┗━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━┛
    /// </summary>
    public abstract class Table
    {
        /// <summary>
        /// 是否横向打印
        /// </summary>
        public bool Landscape { get; set; }
        public Title Title { get; set; }
        public Header Header { get; set; }
        public TableHeader TableHeader { get; set; }
        public TableBody TableBody { get; set; }
        public TableFooter TableFooter { get; set; }
        //public IList<string> Columns { get; set; }
        public Footer Footer { get; set; }
        public int ColumnCount { get; set; }
        public int[] ColumnWidths { get; set; }
        public Table()
        {
            Landscape = false;
        }
        abstract public Row CreateRow();
    }
}
