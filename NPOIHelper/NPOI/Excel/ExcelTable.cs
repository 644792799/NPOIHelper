using NPOIHelper.NPOI.Abstract;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace NPOIHelper.NPOI.Excel
{
    public class ExcelTable : Table
    {
        public ExcelTable()
            : base()
        {
            Title = new ExcelTitle();
            Header = new ExcelHeader();
            TableHeader = new ExcelTableHeader();
            TableBody = new ExcelTableBody();
            TableFooter = new ExcelTableFooter();
            Footer = new ExcelFooter();
        }

        public override Row CreateRow()
        {
            ExcelRow row = new ExcelRow();
            return row;
        }
    }
}
