using NPOIHelper.NPOI.Abstract;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace NPOIHelper.NPOI.Excel
{
    public class ExcelRow : Row
    {
        public ExcelRow()
            : base()
        {

        }
        public override Cell CreateCell()
        {
            Cell cell = new ExcelCell();
            return cell;
        }

        public override void AddCell(Cell cell)
        {
            Cells.Add(cell);
        }
    }
}
