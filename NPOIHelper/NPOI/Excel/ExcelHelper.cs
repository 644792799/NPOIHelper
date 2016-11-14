using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NPOIHelper.NPOI.Abstract;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace NPOIHelper.NPOI.Excel
{
    public class ExcelHelper
    {
        public List<ExcelTable> listTable;
        public ExcelHelper(List<ExcelTable> listTable)
        {
            if (listTable == null)
            {
                throw new Exception("没有需要导出的数据");
            }
            this.listTable = listTable;
        }
        public MemoryStream RenderToXls(bool isOnlyOneSheet = true)
        {
            using (var ms = new MemoryStream())
            {
                HSSFWorkbook workbook = new HSSFWorkbook();
                int rowIndex = 0;
                foreach (var table in listTable)
                {
                    int columnCount = table.ColumnCount;
                    var head = table.Rows.FirstOrDefault(r => r.IsHead);
                    var body = table.Rows.Where(r => !r.IsHead).ToList();
                    HSSFSheet sheet = null;

                    if (rowIndex == 0)
                    {
                        sheet = string.IsNullOrWhiteSpace(table.Title) ? workbook.CreateSheet() as HSSFSheet : workbook.CreateSheet(table.Title) as HSSFSheet;
                        if (table.ColumnWidths != null && table.ColumnWidths.Length == columnCount)
                        {
                            for (var i = 0; i < columnCount; i++)
                            {
                                sheet.SetColumnWidth(i, table.ColumnWidths[i]*256);
                            }
                        }
                        else
                        {
                            if (head != null)
                            {
                                for (var i = 0; i < columnCount; i++)
                                {
                                    int columnWidth = Encoding.Default.GetBytes(head.Cells[i].Value.ToString()).Length;
                                    sheet.SetColumnWidth(i, columnWidth*256);
                                }
                            }
                        }
                        //标题
                        if (!string.IsNullOrWhiteSpace(table.Title))
                        {
                            IRow rowTitle = sheet.CreateRow(rowIndex++);
                            rowTitle.Height = short.Parse(table.TitleHeight * 20 + "");
                            ICell headerCell = rowTitle.CreateCell(0);
                            headerCell.SetCellValue(table.Title);
                            sheet.AddMergedRegion(new CellRangeAddress(0, 0, 0, columnCount - 1));
                            ExcelCellSetter.SetDefaultTitleCellStyle(workbook, headerCell);
                        }
                    }

                    //页头
                    if (table.Header != null)
                    {
                        for (var rindex = 0; rindex < table.Header.Rows.Count; rindex++)
                        {
                            var headerRow = sheet.CreateRow(rowIndex);
                            var row = table.Header.Rows[rindex];
                            headerRow.Height = short.Parse(row.Height * 20 + "");
                            int colindex = 0;
                            for (var cindex = 0; cindex < row.Cells.Count; cindex++)
                            {
                                int colspan = row.Cells[cindex].Colspan;
                                CellRangeAddress address = new CellRangeAddress(rowIndex, rowIndex, colindex, colindex + colspan - 1);
                                sheet.AddMergedRegion(address);
                                headerRow.CreateCell(colindex).SetCellValue(row.Cells[cindex].Value);
                                ExcelCellSetter.SetDefaultHeaderCellStyle(workbook, headerRow.Cells[colindex]);
                                colindex = colindex + colspan;
                            }
                            rowIndex++;
                        }
                    }

                    if (!body.Any())
                    {
                        throw new Exception("没有需要导出的数据");
                    }

                    //表头
                    if (head != null)
                    {
                        var headerRow = sheet.CreateRow(rowIndex++);
                        headerRow.Height = short.Parse(head.Height * 20 + "");
                        for (var i = 0; i < columnCount; i++)
                        {
                            ExcelCell cell = (ExcelCell)head.Cells[i];
                            headerRow.CreateCell(i).SetCellValue(cell.Value);
                            if (cell.CellStyle != null)
                            {
                                ExcelCellSetter.SetCellStyle(headerRow.Cells[i], cell.CellStyle);
                            }
                            else
                            {
                                ExcelCellSetter.SetDefaultTableHeaderCellStyle(workbook, headerRow.Cells[i]);
                            }
                        }
                    }

                    //表体
                    foreach (var row in body)
                    {
                        var dataRow = sheet.CreateRow(rowIndex++);
                        dataRow.Height = short.Parse(row.Height * 20 + "");
                        for (var i = 0; i < columnCount; i++)
                        {
                            dataRow.CreateCell(i, (CellType)row.Cells[i].CellType).SetCellValue(row.Cells[i].Value);
                            ExcelCellSetter.SetDefaultCellStyle(workbook, dataRow.Cells[i]);
                        }
                    }

                    //页尾
                    if (table.Footer != null)
                    {
                        Footer footer = table.Footer;

                    }

                    if (!isOnlyOneSheet)
                    {
                        rowIndex = 0;
                    }
                }

                workbook.Write(ms);
                ms.Position = 0;
                return ms;
            }
        }

        public MemoryStream RenderToXls(ExcelTable table, string sheetName)
        {
            using (var ms = new MemoryStream())
            {

                return null;
            }
        }

        public Stream RenderToXlsx(ExcelTable table)
        {
            //var excel = new ExcelPackage();
            //var workbook = excel.Workbook.Worksheets;
            //var sheet = string.IsNullOrWhiteSpace(table.Title) ? workbook.Add("sheet1") : workbook.Add(table.Title);

            //var head = table.Rows.FirstOrDefault(r => r.IsHead);
            //var body = table.Rows.Where(r => !r.IsHead).ToList();
            //if (!body.Any())
            //{
            //    throw new Exception("没有需要导出的数据");
            //}
            //if (head != null)
            //{
            //    for (var i = 0; i < head.Columns.Count; i++)
            //    {
            //        sheet.Cells[1, i + 1].Value = head.Columns[i].Value;
            //        sheet.Cells[1, i + 1].Style.Font.Name = "微软雅黑";
            //        sheet.Cells[1, i + 1].Style.Font.Color.SetColor(Color.Red);
            //    }
            //    sheet.Row(1).PageBreak = true;

            //    var r = 2;
            //    foreach (var row in body)
            //    {
            //        for (var i = 0; i < row.Columns.Count; i++)
            //        {
            //            sheet.Cells[r, i + 1].Value = row.Columns[i].Value;
            //        }
            //        r++;
            //    }
            //}
            //else
            //{
            //    var r = 2;
            //    foreach (var row in body)
            //    {
            //        for (var i = 0; i < row.Columns.Count; i++)
            //        {
            //            sheet.Cells[r, i + 1].Value = row.Columns[i].Value;
            //        }
            //        r++;
            //    }
            //}
            //if (table.BandColumns.Any())
            //{
            //    foreach (var bandColumn in table.BandColumns)
            //    {
            //        CellSetter.SetCellDropdownlistForXlsx(sheet, bandColumn.Key, bandColumn.Value);
            //    }
            //}

            //excel.Save();
            //excel.Stream.Position = 0;
            //return excel.Stream;
            return null;
        }

        public MemoryStream RenderToXlsx(ExcelTable table, string sheetName)
        {
            using (var ms = new MemoryStream())
            {

                return null;
            }
        }

        public void SaveToFile(MemoryStream ms, string fileName)
        {
            using (var fs = new FileStream(fileName, FileMode.Create, FileAccess.ReadWrite))
            {
                var data = ms.GetBuffer();
                fs.Write(data, 0, data.Length);
                fs.Flush();
            }
        }
        public void SaveToFile(Stream ms, string fileName)
        {
            using (var fs = new FileStream(fileName, FileMode.Create, FileAccess.ReadWrite))
            {
                ms.CopyTo(fs);
                fs.Flush();
            }
        }
    }
}
