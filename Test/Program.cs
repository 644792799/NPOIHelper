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
            List<ExcelTable> l = new List<ExcelTable>();

            for (int k = 0; k < 4; k++)
            {
                ExcelTable table = new ExcelTable();
                table.Title = "疑似黑广播信号出现情况" + (k + 1);
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

                ExcelFooter footer = new ExcelFooter();
                ExcelRow footerrow = (ExcelRow)table.CreateRow();
                footerrow.Height = 200;
                ExcelCell footercell = (ExcelCell)footerrow.CreateCell();
                byte[] b = System.IO.File.ReadAllBytes(@"C:\Users\HRDS-ZENGPEIFENG\Pictures\th2X7ZWG1D.jpg");
                footercell.Value = Convert.ToBase64String(b);
                footercell.Colspan = table.ColumnCount;
                footercell.CellType = NPOIHelper.NPOI.Common.CellTypes.Image;
                footerrow.AddCell(footercell);

                ExcelRow footerrow2 = (ExcelRow)table.CreateRow();
                ExcelCell footercell2 = (ExcelCell)footerrow2.CreateCell();
                footercell2.Colspan = table.ColumnCount;
                footercell2.CellType = NPOIHelper.NPOI.Common.CellTypes.String;
                footercell2.Value = "表格描述信息";
                footerrow2.AddCell(footercell2);

                footercell2 = (ExcelCell)footerrow2.CreateCell();
                footercell2.CellType = NPOIHelper.NPOI.Common.CellTypes.String;
                footercell2.Value = "表格描述信息2";
                footerrow2.AddCell(footercell2);

                List<Row> footerrows = new List<Row>();
                footerrows.Add(new ExcelRow());
                footerrows.Add(footerrow);
                footerrows.Add(footerrow2);
                footer.Rows = footerrows;
                table.Footer = footer;

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

                l.Add(table);
            }
            ExcelHelper excelhelper = new ExcelHelper(l);
            MemoryStream s = excelhelper.RenderToXls(false);
            bool issaved = excelhelper.SaveToFile(s, "d:/test.xls");
            if (issaved)
            {
                excelhelper.ExcelPrint("d:/test.xls", "疑似黑广播信号出现情况1");
            }
        }
    }
}
