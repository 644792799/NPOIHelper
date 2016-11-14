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
        public dynamic Value { get; set; }
        public CellTypes CellType { get; set; }
        //public TextAlign TextAlign { get; set; }
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
        public int Colspan { get; set; }
        public int Rowspan { get; set; }

        public Cell() { }

        public Cell(dynamic value) {
            Value = value;
            CellType = CellTypes.String;
        }
        public Cell(dynamic value,CellTypes cellType) {
            Value = value;
            CellType = cellType;
        }
    }
}
