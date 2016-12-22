using NPOIHelper.NPOI.Abstract;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace NPOIHelper.NPOI.Excel
{
    public class ExcelTableFooter : TableFooter
    {
        public ExcelTableFooter() : base() { }

        public override void AddRow(Row row)
        {
            Rows.Add(row);
        }
    }
}
