﻿using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NPOIHelper.NPOI.Abstract;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Web;

namespace NPOIHelper.NPOI.Excel
{
    public class ExcelHelper
    {
        /// <summary>
        /// 报表数据
        /// </summary>
        public List<ExcelTable> listTable;
        /// <summary>
        /// 构造函数
        /// </summary>
        /// <param name="listTable"></param>
        public ExcelHelper(List<ExcelTable> listTable)
        {
            if (listTable == null)
            {
                throw new Exception("没有需要导出的数据");
            }
            this.listTable = listTable;
        }
        /// <summary>
        /// 渲染到xls
        /// </summary>
        /// <param name="isOnlyOneSheet"></param>
        /// <returns></returns>
        public MemoryStream RenderToXls(bool isOnlyOneSheet = true)
        {
            using (var ms = new MemoryStream())
            {
                HSSFWorkbook workbook = new HSSFWorkbook();
                int rowIndex = 0;
                HSSFSheet sheet = null;
                foreach (var table in listTable)
                {
                    int columnCount = table.ColumnCount;
                    var title = table.Title == null ? new ExcelTitle() : table.Title;
                    var header = table.Header == null ? new ExcelHeader() : table.Header;
                    var tableheader = table.TableHeader == null ? new ExcelTableHeader() : table.TableHeader;//table.Rows.FirstOrDefault(r => r.IsRowHead);
                    var tablebody = table.TableBody == null ? new ExcelTableBody() : table.TableBody;//table.Rows.Where(r => !r.IsRowHead).ToList();
                    var tablefooter = table.TableFooter == null ? new ExcelTableFooter() : table.TableFooter;
                    var footer = table.Footer == null ? new ExcelFooter() : table.Footer;

                    if (rowIndex == 0)
                    {
                        sheet = title.IsNullOrWhiteSpace() ? workbook.CreateSheet() as HSSFSheet : workbook.CreateSheet(title.TableTitle) as HSSFSheet;
                        PrintSetup(sheet, table.Landscape);
                        if (table.ColumnWidths != null && table.ColumnWidths.Length == columnCount)
                        {
                            for (var i = 0; i < columnCount; i++)
                            {
                                sheet.SetColumnWidth(i, table.ColumnWidths[i]*256 + 200);
                            }
                        }
                        else
                        {
                            if (!tableheader.IsNull())
                            {
                                for (var i = 0; i < columnCount; i++)
                                {
                                    int columnWidth = Encoding.Default.GetBytes(tableheader.Rows[0].Cells[i].Value.ToString()).Length;
                                    sheet.SetColumnWidth(i, columnWidth * 256 + 200);
                                }
                            }
                        }
                    }

                    //标题
                    if (!title.IsNullOrWhiteSpace())
                    {
                        IRow rowTitle = sheet.CreateRow(rowIndex);
                        rowTitle.Height = short.Parse(title.TitleHeight * 20 + "");
                        ICell headerCell = rowTitle.CreateCell(0);
                        headerCell.SetCellValue(title.TableTitle);
                        sheet.AddMergedRegion(new CellRangeAddress(rowIndex, rowIndex, 0, columnCount - 1));
                        var defaultTitleCellStyle = string.IsNullOrWhiteSpace(title.DefaultTitleCellStyle) ?
                            workbook.GetDefaultTitleCellStyle() : workbook.GetUserDirCellStyle(title.DefaultTitleCellStyle);
                        headerCell.Style(defaultTitleCellStyle);
                        //ExcelCellSetter.SetDefaultTitleCellStyle(workbook, headerCell);
                        rowIndex++;
                    }

                    //页头
                    if (!header.IsNull())
                    {
                        ICellStyle defaultHeaderCellStyle = string.IsNullOrWhiteSpace(header.DefaultHeaderCellStyle) ?
                            workbook.GetDefaultHeaderCellStyle() : workbook.GetUserDirCellStyle(header.DefaultHeaderCellStyle);//ExcelCellSetter.GetDefaultHeaderCellStyle(workbook);
                        for (var rindex = 0; rindex < header.Rows.Count; rindex++)
                        {
                            var headerRow = sheet.CreateRow(rowIndex);
                            var row = header.Rows[rindex];
                            if (row.HaveRowBreak)
                            {
                                SetRowBreak(rowIndex, sheet);
                            }
                            headerRow.Height = short.Parse(row.Height * 20 + "");
                            int colindex = 0;
                            for (var cindex = 0; cindex < row.Cells.Count; cindex++)
                            {
                                int colspan = row.Cells[cindex].Colspan;
                                CellRangeAddress address = new CellRangeAddress(rowIndex, rowIndex, colindex, colindex + colspan - 1);
                                sheet.AddMergedRegion(address);
                                headerRow.CreateCell(colindex).SetCellValue(row.Cells[cindex].Value);

                                if (row.Cells[cindex].CellStyle != null)
                                {
                                    headerRow.Cells[colindex].Style(row.Cells[cindex].CellStyle);
                                    //ExcelCellSetter.SetCellStyle(headerRow.Cells[colindex], row.Cells[cindex].CellStyle);
                                }
                                else if (row.Cells[cindex].Style != null)
                                {
                                    if (row.Cells[cindex].IsBasedOnDefaultStyle)
                                    {
                                        headerRow.Cells[colindex].Style(defaultHeaderCellStyle);
                                    }
                                    headerRow.Cells[colindex].Style(row.Cells[cindex].Style);
                                }
                                else
                                {
                                    headerRow.Cells[colindex].Style(defaultHeaderCellStyle);
                                    //ExcelCellSetter.SetCellStyle(headerRow.Cells[colindex], defaultHeaderCellStyle);
                                }
                                colindex = colindex + colspan;
                            }
                            rowIndex++;
                        }
                    }

                    //表头
                    if (!tableheader.IsNull())
                    {
                        ICellStyle defaultTableHeaderCellStyle =
                            string.IsNullOrWhiteSpace(
                            tableheader.DefaultTableHeaderCellStyle) ?
                            workbook.GetDefaultTableHeaderCellStyle() :
                            workbook.GetUserDirCellStyle(tableheader.DefaultTableHeaderCellStyle);//ExcelCellSetter.GetDefaultTableHeaderCellStyle(workbook);
                        foreach (var hr in tableheader.Rows)
                        {
                            var headerRow = sheet.CreateRow(rowIndex++);
                            headerRow.Height = short.Parse(hr.Height * 20 + "");
                            for (var i = 0; i < columnCount; i++)
                            {
                                ExcelCell cell = (ExcelCell)hr.Cells[i];
                                headerRow.CreateCell(i).SetCellValue(cell.Value);
                                if (cell.CellStyle != null)
                                {
                                    headerRow.Cells[i].Style(cell.CellStyle);
                                    //ExcelCellSetter.SetCellStyle(headerRow.Cells[i], cell.CellStyle);
                                }
                                else if (cell.Style != null)
                                {
                                    if (cell.IsBasedOnDefaultStyle)
                                    {
                                        headerRow.Cells[i].Style(defaultTableHeaderCellStyle);
                                    }
                                    headerRow.Cells[i].Style(cell.Style);
                                }
                                else
                                {
                                    headerRow.Cells[i].Style(defaultTableHeaderCellStyle);
                                    //ExcelCellSetter.SetCellStyle(headerRow.Cells[i], defaultTableHeaderCellStyle);
                                }
                            }
                        }
                    }

                    //表体
                    if (!tablebody.IsNull())
                    {
                        ICellStyle defaultcellstyle = string.IsNullOrWhiteSpace(tablebody.DefaultTableBodyCellStyle) ?
                            workbook.GetDefaultTableBodyCellStyle() : workbook.GetUserDirCellStyle(tablebody.DefaultTableBodyCellStyle);//ExcelCellSetter.GetDefaultCellStyle(workbook);
                        //表体
                        foreach (var row in tablebody.Rows)
                        {
                            var dataRow = sheet.CreateRow(rowIndex++);
                            if (row.HaveRowBreak)
                            {
                                SetRowBreak(rowIndex, sheet);
                            }
                            dataRow.Height = short.Parse(row.Height * 20 + "");
                            for (var i = 0; i < columnCount; i++)
                            {
                                ExcelCell cell = (ExcelCell)row.Cells[i];
                                var cellval = cell.Value == null ? "" : cell.Value;
                                dataRow.CreateCell(i, (CellType)cell.CellType).SetCellValue(cellval);
                                if (cell.CellStyle != null)
                                {
                                    dataRow.Cells[i].Style(cell.CellStyle);
                                    //ExcelCellSetter.SetCellStyle(dataRow.Cells[i], cell.CellStyle);
                                }
                                else if (cell.Style != null)
                                {
                                    if (cell.IsBasedOnDefaultStyle)
                                    {
                                        dataRow.Cells[i].Style(defaultcellstyle);
                                    }
                                    dataRow.Cells[i].Style(cell.Style);
                                }
                                else
                                {
                                    dataRow.Cells[i].Style(defaultcellstyle);
                                    //ExcelCellSetter.SetCellStyle(dataRow.Cells[i], defaultcellstyle);
                                }
                            }
                        }
                    }

                    //表尾
                    if (!tablefooter.IsNull())
                    {
                        ICellStyle defaultTableFooterCellStyle =
                            string.IsNullOrWhiteSpace(
                            tablefooter.DefaultTableFooterCellStyle) ?
                            workbook.GetDefaultTableFooterCellStyle() :
                            workbook.GetUserDirCellStyle(tablefooter.DefaultTableFooterCellStyle);//ExcelCellSetter.GetDefaultTableHeaderCellStyle(workbook);
                        foreach (var hr in tablefooter.Rows)
                        {
                            var footerRow = sheet.CreateRow(rowIndex++);
                            footerRow.Height = short.Parse(hr.Height * 20 + "");
                            for (var i = 0; i < columnCount; i++)
                            {
                                ExcelCell cell = (ExcelCell)hr.Cells[i];
                                footerRow.CreateCell(i).SetCellValue(cell.Value);
                                if (cell.CellStyle != null)
                                {
                                    footerRow.Cells[i].Style(cell.CellStyle);
                                }
                                else if (cell.Style != null)
                                {
                                    if (cell.IsBasedOnDefaultStyle)
                                    {
                                        footerRow.Cells[i].Style(defaultTableFooterCellStyle);
                                    }
                                    footerRow.Cells[i].Style(cell.Style);
                                }
                                else
                                {
                                    footerRow.Cells[i].Style(defaultTableFooterCellStyle);
                                }
                            }
                        }
                    }

                    //页尾
                    if (!footer.IsNull())
                    {
                        ICellStyle defaultFooterCellStyle = string.IsNullOrWhiteSpace(footer.DefaultFooterCellStyle) ?
                            workbook.GetDefaultFooterCellStyle() : workbook.GetUserDirCellStyle(footer.DefaultFooterCellStyle);//ExcelCellSetter.GetDefaultFooterCellStyle(workbook);
                        for (var rindex = 0; rindex < footer.Rows.Count; rindex++)
                        {
                            var footerRow = sheet.CreateRow(rowIndex);
                            var row = footer.Rows[rindex];
                            if (row.HaveRowBreak)
                            {
                                SetRowBreak(rowIndex, sheet);
                            }
                            //1 px = 0.75 point
                            footerRow.Height = short.Parse(row.Height * 0.75 * 20 + "");
                            int colindex = 0;
                            for (var cindex = 0; cindex < row.Cells.Count; cindex++)
                            {
                                int colspan = row.Cells[cindex].Colspan;
                                CellRangeAddress address = new CellRangeAddress(rowIndex, rowIndex, colindex, colindex + colspan - 1);
                                sheet.AddMergedRegion(address);
                                if (row.Cells[cindex].CellType == Common.CellTypes.Image)
                                {
                                    footerRow.CreateCell(colindex);
                                    byte[] imgbyte = Convert.FromBase64String(row.Cells[cindex].Value);
                                    int pictureIdx = workbook.AddPicture(imgbyte, PictureType.PNG);
                                    HSSFPatriarch patriarch = (HSSFPatriarch)sheet.CreateDrawingPatriarch();
                                    HSSFClientAnchor anchor = new HSSFClientAnchor(0, 1, 0, 0, 0, rowIndex, colspan, rowIndex + 1);
                                    anchor.AnchorType = AnchorType.DontMoveAndResize;
                                    HSSFPicture pict = (HSSFPicture)patriarch.CreatePicture(anchor, pictureIdx);
                                    pict.Resize(0.99, 0.99);
                                    //pict.Resize();
                                }
                                else
                                {
                                    footerRow.CreateCell(colindex).SetCellValue(row.Cells[cindex].Value);
                                }

                                if (row.Cells[cindex].CellStyle != null)
                                {
                                    footerRow.Cells[colindex].Style(row.Cells[cindex].CellStyle);
                                    //ExcelCellSetter.SetCellStyle(footerRow.Cells[colindex], row.Cells[cindex].CellStyle);
                                }
                                else if (row.Cells[cindex].Style != null)
                                {
                                    if (row.Cells[cindex].IsBasedOnDefaultStyle)
                                    {
                                        footerRow.Cells[colindex].Style(defaultFooterCellStyle);
                                    }
                                    footerRow.Cells[colindex].Style(row.Cells[cindex].Style);
                                }
                                else
                                {
                                    footerRow.Cells[colindex].Style(defaultFooterCellStyle);
                                    //if (defaultFooterCellStyle != null)
                                        //ExcelCellSetter.SetCellStyle(footerRow.Cells[colindex], defaultFooterCellStyle);
                                }
                                colindex = colindex + colspan;
                            }
                            rowIndex++;
                        }
                    }

                    if (!isOnlyOneSheet)
                    {
                        rowIndex = 0;
                    }
                }

                workbook.Write(ms);
                ms.Position = 0;

                workbook.Close();
                return ms;
            }
        }
        /// <summary>
        /// 打印设置
        /// 分辨率是96像素/英寸时，A4纸的尺寸的图像的像素是794×1123
        /// </summary>
        /// <param name="sheet"></param>
        private void PrintSetup(HSSFSheet sheet, bool landscape)
        {
            sheet.PrintSetup.PaperSize = (short)9;//A4
            sheet.PrintSetup.Landscape = landscape;// true 横向打印 false 竖向打印
            sheet.PrintSetup.Scale = 95;
            sheet.SetMargin(MarginType.RightMargin, (double)0.1);
            sheet.SetMargin(MarginType.TopMargin, (double)0.1);
            sheet.SetMargin(MarginType.LeftMargin, (double)0.1);
            sheet.SetMargin(MarginType.BottomMargin, (double)0.1);
            sheet.HorizontallyCenter = true;

            sheet.FitToPage = false;
            //不改变fittopage属性情况下实现打印分页效果设置
            sheet.PrintSetup.FitHeight = 1;
            sheet.PrintSetup.FitWidth = 1; // this is the default value
        }
        /// <summary>
        /// 设置打印分页符
        /// </summary>
        /// <param name="rowIndex"></param>
        /// <param name="sheet"></param>
        private void SetRowBreak(int rowIndex, HSSFSheet sheet)
        {
            sheet.SetRowBreak(rowIndex);
        }
        /// <summary>
        /// 渲染到xls
        /// </summary>
        /// <param name="table"></param>
        /// <param name="sheetName"></param>
        /// <returns></returns>
        public MemoryStream RenderToXls(ExcelTable table, string sheetName)
        {
            using (var ms = new MemoryStream())
            {

                return null;
            }
        }
        /// <summary>
        /// 渲染到xlsx
        /// </summary>
        /// <param name="table"></param>
        /// <returns></returns>
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
        /// <summary>
        /// 渲染到xlsx
        /// </summary>
        /// <param name="table"></param>
        /// <param name="sheetName"></param>
        /// <returns></returns>
        public MemoryStream RenderToXlsx(ExcelTable table, string sheetName)
        {
            using (var ms = new MemoryStream())
            {

                return null;
            }
        }
        /// <summary>
        /// 保存文件
        /// </summary>
        /// <param name="ms"></param>
        /// <param name="fileName"></param>
        /// <returns></returns>
        public bool SaveToFile(MemoryStream ms, string fileName)
        {
            using (var fs = new FileStream(fileName, FileMode.Create, FileAccess.ReadWrite))
            {
                var data = ms.GetBuffer();
                fs.Write(data, 0, data.Length);
                fs.Flush();
                fs.Close();
                return true;
            }
        }
        /// <summary>
        /// 保存文件
        /// </summary>
        /// <param name="ms"></param>
        /// <param name="fileName"></param>
        /// <returns></returns>
        public bool SaveToFile(Stream ms, string fileName)
        {
            using (var fs = new FileStream(fileName, FileMode.Create, FileAccess.ReadWrite))
            {
                ms.CopyTo(fs);
                fs.Flush();
                fs.Close();
                return true;
            }
        }
        /// <summary>
        /// 打印excel（需要安装office）
        /// </summary>
        /// <param name="strFilePath"></param>
        /// <param name="strSheetName"></param>
        public void ExcelPrint(string strFilePath, string strSheetName)
        {
            //Microsoft.Office.Interop.Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();
            //Microsoft.Office.Interop.Excel.Workbook xlWorkbook;
            //Microsoft.Office.Interop.Excel.Worksheet xlWorksheet;

            //object oMissing = System.Reflection.Missing.Value;
            ////strFilePath = Server.MapPath(strFilePath);
            //if (!System.IO.File.Exists(strFilePath))
            //{
            //    throw new System.IO.FileNotFoundException();
            //    return;
            //}
            //try
            //{
            //    xlApp.Visible = true;
            //    xlWorkbook = xlApp.Workbooks.Add(strFilePath);
            //    xlWorksheet = (Microsoft.Office.Interop.Excel.Worksheet)xlWorkbook.Worksheets[strSheetName];
            //    xlWorksheet.PrintPreview(false);
            //    xlWorkbook.Close(oMissing, oMissing, oMissing);
            //}
            //catch (Exception ex)
            //{
            //    throw ex;
            //}
            //finally
            //{
            //    if (xlApp != null)
            //    {
            //        xlApp.Quit();
            //        System.Runtime.InteropServices.Marshal.ReleaseComObject(xlApp);
            //        xlApp = null;
            //    }
            //    GC.Collect();
            //}
        }
    }
}
