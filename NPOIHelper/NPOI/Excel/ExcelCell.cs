using NPOIHelper.NPOI.Abstract;
using NPOIHelper.NPOI.Common;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace NPOIHelper.NPOI.Excel
{
    public class ExcelCell : Cell
    {
        /// <summary>
        /// 单元格批注
        /// </summary>
        public string Comments { get; set; }

        public ExcelCell()
            : base()
        {
            CellType = CellTypes.String;
            Colspan = 1;
            Rowspan = 1;
        }

        public ExcelCell(dynamic value)
        {
            Value = value;
            CellType = CellTypes.String;
            Colspan = 1;
            Rowspan = 1;
        }

        public ExcelCell(dynamic value, CellTypes cellType)
        {
            Value = value;
            CellType = cellType;
            Colspan = 1;
            Rowspan = 1;
        }
    }
}
