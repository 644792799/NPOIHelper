﻿using NPOIHelper.Contract;
using NPOIHelper.NPOI.Abstract;
using NPOIHelper.NPOI.Excel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.Remoting;
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
                table.Landscape = true;
                //table.Title = "疑似黑广播信号出现情况" + (k + 1);
                table.ColumnCount = 10;
                int[] columnswidth = new int[table.ColumnCount];

                ExcelTitle title = new ExcelTitle();
                title.TableTitle = "疑似黑广播信号出现情况" + (k + 1);
                table.Title = title;

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
                //ExcelRow footerrow = (ExcelRow)table.CreateRow();
                //footerrow.Height = 200;
                //ExcelCell footercell = (ExcelCell)footerrow.CreateCell();
                //byte[] b = System.IO.File.ReadAllBytes(@"C:\Users\HRDS-ZENGPEIFENG\Pictures\th2X7ZWG1D.jpg");
                //footercell.Value = Convert.ToBase64String(b);
                //footercell.Colspan = table.ColumnCount;
                //footercell.CellType = NPOIHelper.NPOI.Common.CellTypes.Image;
                //footerrow.AddCell(footercell);

                ExcelRow footerrow2 = (ExcelRow)table.CreateRow();
                ExcelCell footercell2 = (ExcelCell)footerrow2.CreateCell();
                footercell2.Colspan = table.ColumnCount;
                footercell2.CellType = NPOIHelper.NPOI.Common.CellTypes.String;
                footercell2.Value = "表格描述信息";
                footerrow2.AddCell(footercell2);

                //footercell2 = (ExcelCell)footerrow2.CreateCell();
                //footercell2.CellType = NPOIHelper.NPOI.Common.CellTypes.String;
                //footercell2.Value = "表格描述信息2";
                //footerrow2.AddCell(footercell2);
                footerrow2.HaveRowBreak = true;

                List<Row> footerrows = new List<Row>();
                footerrows.Add(new ExcelRow());
                //footerrows.Add(footerrow);
                footerrows.Add(footerrow2);
                footer.Rows = footerrows;
                table.Footer = footer;

                ExcelTableHeader tableheader = new ExcelTableHeader();
                ExcelTableBody tablebody = new ExcelTableBody();

                //throws.Add(new )

                for (int r = 0; r < table.ColumnCount; r++)
                {
                    ExcelRow row;
                    if (r == 0 || r == 1)
                    {
                        row = (ExcelRow)table.CreateRow();
                        row.Height = 28;
                    }
                    else
                    {
                        row = (ExcelRow)table.CreateRow();
                    }

                    for (int i = 0; i < table.ColumnCount; i++)
                    {
                        ExcelCell cell = (ExcelCell)row.CreateCell();
                        if (r == 4 && i == 5)
                        {
                            cell.IsBasedOnDefaultStyle = true;
                            cell.Style = "font-color:#03A8F4;border-color:#03A8F4;font-size:50;";
                        }
                        //if (i == 1)
                        //{
                        //    cell.CellType = NPOIHelper.NPOI.Common.CellTypes.Numeric;
                        //    cell.Value = r * i;
                        //    cell.Style = "border:thin;font-color:red;font-weight:normal;text-align:left;";
                        //}
                        //else
                        //{
                            cell.CellType = NPOIHelper.NPOI.Common.CellTypes.String;
                            if (r == 0 || r == 1)
                            {
                                cell.Value = "标题" + i + "\r\n(单位)";
                                columnswidth[i] = 12;
                            }
                            else
                            {
                                cell.Value = "行:" + r + " 列:" + i;
                            }
                        //}

                        //cell.FontColor = HSSFColor.BlueGrey.Index;
                        row.AddCell(cell);
                        
                    }
                    if (r == 0 || r == 1)
                    {
                        tableheader.AddRow(row);
                    }
                    else
                    {
                        tablebody.AddRow(row);
                    }
                }
                table.TableHeader = tableheader;
                table.TableBody = tablebody;
                table.ColumnWidths = columnswidth;
                l.Add(table);
            }
            ExcelHelper excelhelper = new ExcelHelper(l);
            MemoryStream s = excelhelper.RenderToXls();
            bool issaved = excelhelper.SaveToFile(s, "d:/test.xls");
            //issaved = false;
            if (issaved)
            {
                IPrint proxy = null;
                RemotingConfiguration.Configure(AppDomain.CurrentDomain.SetupInformation.ConfigurationFile, false);
                proxy = (IPrint)Activator.GetObject(typeof(IPrint), "tcp://172.39.8.173:1235/Print/PrintURL");
                proxy.ExcelPrint("d:/test.xls", "疑似黑广播信号出现情况1", new PrintCallBackHandler());

                //excelhelper.ExcelPrint("d:/test.xls", "疑似黑广播信号出现情况1");
            }
        }
    }
}
