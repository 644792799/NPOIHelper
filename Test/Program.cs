using NPOI.HSSF.Util;
using NPOIHelper.NPOI.Abstract;
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
            int[] columnswidth = new int[table.ColumnCount];

            ExcelHeader header = new ExcelHeader();
            
            ExcelRow headerrow = (ExcelRow)table.CreateRow();
            headerrow.Height = 20;
            ExcelCell headercell = (ExcelCell)headerrow.CreateCell();
            headercell.Value = "表格描述信息";
            headercell.Colspan = table.ColumnCount;
            headerrow.AddCell(headercell);

            List<Row> rows = new List<Row>();
            rows.Add(new ExcelRow());
            rows.Add(headerrow);
            header.Rows = rows;
            table.Header = header;

            for (int r = 0; r < 5; r++)
            {
                ExcelRow row;
                if (r == 0)
                {
                    row = (ExcelRow)table.CreateRow(true);
                    row.Height = 28;
                }
                else
                {
                    row = (ExcelRow)table.CreateRow();
                }

                for (int i = 0; i < table.ColumnCount; i++)
                {
                    ExcelCell cell = (ExcelCell)row.CreateCell();
                    if (i == 1)
                    {
                        cell.CellType = NPOIHelper.NPOI.Common.CellTypes.Numeric;
                        cell.Value = r * i;
                    }
                    else
                    {
                        cell.CellType = NPOIHelper.NPOI.Common.CellTypes.String;
                        cell.Value = "行:" + r + " 列:" + i;
                    }
                    
                    //cell.FontColor = HSSFColor.BlueGrey.Index;
                    row.AddCell(cell);
                    columnswidth[i] = 20;
                }
                table.ColumnWidths = columnswidth;
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
