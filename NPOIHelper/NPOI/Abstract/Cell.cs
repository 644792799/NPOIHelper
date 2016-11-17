using NPOI.SS.UserModel;
using NPOIHelper.NPOI.Common;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace NPOIHelper.NPOI.Abstract
{
    public abstract class Cell
    {
        /// <summary>
        /// 单元格值
        /// </summary>
        public dynamic Value { get; set; }
        /// <summary>
        /// 单元格类型
        /// </summary>
        public CellTypes CellType { get; set; }
        /// <summary>
        /// 单元格样式
        /// </summary>
        private ICellStyle cellStyle;
        public ICellStyle CellStyle
        {
            get
            {
                return cellStyle;
            }
            set
            {
                this.cellStyle = value;
            }
        }
        /// <summary>
        /// 跨列
        /// </summary>
        public int Colspan { get; set; }
        /// <summary>
        /// 跨行
        /// </summary>
        public int Rowspan { get; set; }

        public Cell()
        {
            CellType = CellTypes.String;
            Colspan = 1;
            Rowspan = 1;
        }

        public Cell(dynamic value) {
            Value = value;
            CellType = CellTypes.String;
            Colspan = 1;
            Rowspan = 1;
        }

        public Cell(dynamic value,CellTypes cellType) {
            Value = value;
            CellType = cellType;
            Colspan = 1;
            Rowspan = 1;
        }
    }
}
