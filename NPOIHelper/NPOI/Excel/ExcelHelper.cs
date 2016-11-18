using Microsoft.Office.Interop.Excel;
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
                    var head = table.Rows.FirstOrDefault(r => r.IsRowHead);
                    var body = table.Rows.Where(r => !r.IsRowHead).ToList();

                    if (rowIndex == 0)
                    {
                        sheet = string.IsNullOrWhiteSpace(table.Title) ? workbook.CreateSheet() as HSSFSheet : workbook.CreateSheet(table.Title) as HSSFSheet;
                        PrintSetup(sheet);
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
                            ICellStyle defaultHeaderCellStyle = ExcelCellSetter.GetDefaultHeaderCellStyle(workbook);
                            for (var cindex = 0; cindex < row.Cells.Count; cindex++)
                            {
                                int colspan = row.Cells[cindex].Colspan;
                                CellRangeAddress address = new CellRangeAddress(rowIndex, rowIndex, colindex, colindex + colspan - 1);
                                sheet.AddMergedRegion(address);
                                headerRow.CreateCell(colindex).SetCellValue(row.Cells[cindex].Value);

                                if (row.Cells[cindex].CellStyle != null)
                                {
                                    ExcelCellSetter.SetCellStyle(headerRow.Cells[colindex], row.Cells[cindex].CellStyle);
                                }
                                else
                                {
                                    ExcelCellSetter.SetCellStyle(headerRow.Cells[colindex], defaultHeaderCellStyle);
                                    //ExcelCellSetter.SetDefaultHeaderCellStyle(workbook, headerRow.Cells[colindex]);
                                }
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
                        ICellStyle defaultTableHeaderCellStyle = ExcelCellSetter.GetDefaultTableHeaderCellStyle(workbook);
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
                                ExcelCellSetter.SetCellStyle(headerRow.Cells[i], defaultTableHeaderCellStyle);
                                //ExcelCellSetter.SetDefaultTableHeaderCellStyle(workbook, headerRow.Cells[i]);
                            }
                        }
                    }

                    //表体
                    foreach (var row in body)
                    {
                        var dataRow = sheet.CreateRow(rowIndex++);
                        dataRow.Height = short.Parse(row.Height * 20 + "");
                        ICellStyle defaultcellstyle = ExcelCellSetter.GetDefaultCellStyle(workbook);
                        for (var i = 0; i < columnCount; i++)
                        {
                            ExcelCell cell = (ExcelCell)row.Cells[i];
                            var cellval = cell.Value == null ? "" : cell.Value;
                            dataRow.CreateCell(i, (CellType)cell.CellType).SetCellValue(cellval);
                            if (cell.CellStyle != null)
                            {
                                ExcelCellSetter.SetCellStyle(dataRow.Cells[i], cell.CellStyle);
                            }
                            else
                            {
                                //ExcelCellSetter.SetDefaultCellStyle(workbook, dataRow.Cells[i]);
                                ExcelCellSetter.SetCellStyle(dataRow.Cells[i], defaultcellstyle);
                            }
                        }
                    }

                    //页尾
                    if (table.Footer != null)
                    {
                        Footer footer = table.Footer;
                        for (var rindex = 0; rindex < table.Footer.Rows.Count; rindex++)
                        {
                            var footerRow = sheet.CreateRow(rowIndex);
                            var row = table.Footer.Rows[rindex];
                            footerRow.Height = short.Parse(row.Height * 20 + "");
                            int colindex = 0;
                            ICellStyle defaultFooterCellStyle = ExcelCellSetter.GetDefaultFooterCellStyle(workbook);
                            for (var cindex = 0; cindex < row.Cells.Count; cindex++)
                            {
                                int colspan = row.Cells[cindex].Colspan;
                                CellRangeAddress address = new CellRangeAddress(rowIndex, rowIndex, colindex, colindex + colspan - 1);
                                sheet.AddMergedRegion(address);
                                if (row.Cells[cindex].CellType == Common.CellTypes.Image)
                                {
                                    byte[] imgbyte = Convert.FromBase64String(row.Cells[cindex].Value);
                                    int pictureIdx = workbook.AddPicture(imgbyte, PictureType.PNG);
                                    HSSFPatriarch patriarch = (HSSFPatriarch)sheet.CreateDrawingPatriarch();
                                    HSSFClientAnchor anchor = new HSSFClientAnchor(0, 1, 0, 0, 0, rowIndex, colspan, rowIndex + 1);
                                    HSSFPicture pict = (HSSFPicture)patriarch.CreatePicture(anchor, pictureIdx);
                                }
                                else
                                {
                                    footerRow.CreateCell(colindex).SetCellValue(row.Cells[cindex].Value);
                                }

                                if (row.Cells[cindex].CellStyle != null)
                                {
                                    ExcelCellSetter.SetCellStyle(footerRow.Cells[colindex], row.Cells[cindex].CellStyle);
                                }
                                else
                                {
                                    if (defaultFooterCellStyle != null)
                                        ExcelCellSetter.SetCellStyle(footerRow.Cells[colindex], defaultFooterCellStyle);
                                    //ExcelCellSetter.SetDefaultHeaderCellStyle(workbook, headerRow.Cells[colindex]);
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
        /// </summary>
        /// <param name="sheet"></param>
        private void PrintSetup(HSSFSheet sheet)
        {
            sheet.PrintSetup.PaperSize = 9;//A4
            sheet.PrintSetup.Landscape = false;//横向打印
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
            Microsoft.Office.Interop.Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel.Workbook xlWorkbook;
            Microsoft.Office.Interop.Excel.Worksheet xlWorksheet;

            object oMissing = System.Reflection.Missing.Value;
            //strFilePath = Server.MapPath(strFilePath);
            if (!System.IO.File.Exists(strFilePath))
            {
                throw new System.IO.FileNotFoundException();
                return;
            }
            try
            {
                xlApp.Visible = true;
                xlWorkbook = xlApp.Workbooks.Add(strFilePath);
                xlWorksheet = (Microsoft.Office.Interop.Excel.Worksheet)xlWorkbook.Worksheets[strSheetName];
                xlWorksheet.PrintPreview(false);
                xlWorkbook.Close(oMissing, oMissing, oMissing);
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                if (xlApp != null)
                {
                    xlApp.Quit();
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(xlApp);
                    xlApp = null;
                }
                GC.Collect();
            }
        }
    }
}
