using NPOI.HSSF.Util;
using NPOIHelper.NPOI.Excel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Test
{
    class Program
    {
        static void Main(string[] args)
        {
            ExcelTable table = new ExcelTable();
            table.Title = "测试";
            table.ColumnCount = 5;

            for (int r = 0; r < 5; r++)
            {
                ExcelRow row;
                if (r == 0)
                {
                    row = (ExcelRow)table.CreateRow(true);
                }
                else
                {
                    row = (ExcelRow)table.CreateRow();
                }

                for (int i = 0; i < table.ColumnCount; i++)
                {
                    ExcelCell cell = (ExcelCell)row.CreateCell();
                    cell.CellType = NPOIHelper.NPOI.Common.CellTypes.String;
                    cell.Value = "行:" + r + " 列:" + i;
                    //cell.FontColor = HSSFColor.BlueGrey.Index;
                    row.AddCell(cell);
                }
                table.AddRow(row);
            }
            List<ExcelTable> l = new List<ExcelTable>();
            l.Add(table);
            ExcelHelper excelhelper = new ExcelHelper(l);
            MemoryStream s = excelhelper.RenderToXls();
            excelhelper.SaveToFile(s, "d:/test.xls");
        }
    }
}
