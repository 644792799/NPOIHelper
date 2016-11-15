using NPOI.HSSF.UserModel;
using NPOI.HSSF.Util;
using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace NPOIHelper.NPOI.Excel
{
    public class ExcelCellSetter
    {
        /// <summary>
        /// 设置单元格样式
        /// </summary>
        /// <param name="cell"></param>
        /// <param name="cellstyle"></param>
        public static void SetCellStyle(ICell cell, ICellStyle cellstyle)
        {
            cell.CellStyle = cellstyle;
        }

        /// <summary>
        /// 获取默认单元格样式
        /// </summary>
        /// <param name="wb"></param>
        /// <returns></returns>
        public static ICellStyle GetDefaultCellStyle(IWorkbook wb)
        {
            ICellStyle cellStyle = wb.CreateCellStyle();
            IFont ifont = wb.CreateFont();
            ifont.FontName = "宋体";
            ifont.FontHeightInPoints = 9;
            ifont.Boldweight = (short)FontBoldWeight.Normal;
            ifont.Color = HSSFColor.Grey50Percent.Index;
            cellStyle.BorderLeft = BorderStyle.Thin;
            cellStyle.BorderRight = BorderStyle.Thin;
            cellStyle.BorderTop = BorderStyle.Thin;
            cellStyle.BorderBottom = BorderStyle.Thin;
            cellStyle.LeftBorderColor = HSSFColor.Black.Index;
            return cellStyle;
        }

        /// <summary>
        /// 设置默认单元格样式
        /// </summary>
        /// <param name="workbook"></param>
        /// <param name="cell"></param>
        public static void SetDefaultCellStyle(IWorkbook workbook, ICell cell)
        {
            ICellStyle cellStyle = GetDefaultCellStyle(workbook);
            cell.CellStyle = cellStyle;
        }

        /// <summary>
        /// 设置默认标题样式
        /// </summary>
        /// <param name="workbook"></param>
        /// <param name="cell"></param>
        public static void SetDefaultTitleCellStyle(IWorkbook workbook, ICell cell)
        {
            ICellStyle cellStyle = GetDefaultCellStyle(workbook);
            //边框  
            cellStyle.BorderLeft = BorderStyle.None;
            cellStyle.BorderRight = BorderStyle.None;
            cellStyle.BorderTop = BorderStyle.None;
            cellStyle.BorderBottom = BorderStyle.None;

            IFont ifont = workbook.CreateFont();//cellStyle.GetFont(workbook);
            
            cellStyle.VerticalAlignment = VerticalAlignment.Center;
            cellStyle.Alignment = HorizontalAlignment.Center;
            cellStyle.WrapText = true;
            ifont.FontName = "方正姚体";
            ifont.FontHeightInPoints = 20;
            ifont.Boldweight = (short)FontBoldWeight.Bold;
            ifont.Color = HSSFColor.Grey50Percent.Index;
            cellStyle.SetFont(ifont);

            cell.CellStyle = cellStyle;
        }

        /// <summary>
        /// 获取默认表头样式
        /// </summary>
        /// <param name="workbook"></param>
        /// <returns></returns>
        public static ICellStyle GetDefaultTableHeaderCellStyle(IWorkbook workbook)
        {
            ICellStyle cellStyle = GetDefaultCellStyle(workbook);
            IFont ifont = workbook.CreateFont();//cellStyle.GetFont(workbook);
            //背景
            cellStyle.FillPattern = FillPattern.SolidForeground;
            cellStyle.FillForegroundColor = HSSFColor.Grey25Percent.Index;
            //水平对齐 
            cellStyle.Alignment = HorizontalAlignment.Center;
            //垂直对齐
            cellStyle.VerticalAlignment = VerticalAlignment.Center;

            //字体样式
            ifont.FontName = "方正姚体";
            ifont.FontHeightInPoints = 12;
            ifont.Boldweight = (short)FontBoldWeight.Bold;
            ifont.Color = HSSFColor.Black.Index;
            cellStyle.SetFont(ifont);

            return cellStyle;
        }

        /// <summary>
        /// 设置默认表头样式
        /// </summary>
        /// <param name="workbook"></param>
        /// <param name="cell"></param>
        public static void SetDefaultTableHeaderCellStyle(IWorkbook workbook, ICell cell)
        {
            ICellStyle cellStyle = GetDefaultTableHeaderCellStyle(workbook);
            cell.CellStyle = cellStyle;
        }

        /// <summary>
        /// 设置默认表体样式
        /// </summary>
        /// <param name="workbook"></param>
        /// <param name="cell"></param>
        public static void SetDefaultTableBodyCellStyle(IWorkbook workbook, ICell cell)
        {

        }

        /// <summary>
        /// 设置默认表尾样式
        /// </summary>
        /// <param name="workbook"></param>
        /// <param name="cell"></param>
        public static void SetDefaultTableFooterCellStyle(IWorkbook workbook, ICell cell)
        {

        }

        /// <summary>
        /// 获取默认页头样式
        /// </summary>
        /// <param name="workbook"></param>
        /// <returns></returns>
        public static ICellStyle GetDefaultHeaderCellStyle(IWorkbook workbook)
        {
            ICellStyle cellStyle = GetDefaultCellStyle(workbook);
            //边框  
            cellStyle.BorderLeft = BorderStyle.None;
            cellStyle.BorderRight = BorderStyle.None;
            cellStyle.BorderTop = BorderStyle.None;
            cellStyle.BorderBottom = BorderStyle.None;

            IFont ifont = workbook.CreateFont();//cellStyle.GetFont(workbook);

            //水平对齐 
            cellStyle.Alignment = HorizontalAlignment.Left;
            //垂直对齐
            cellStyle.VerticalAlignment = VerticalAlignment.Center;
            cellStyle.WrapText = true;

            //字体样式
            ifont.FontName = "宋体";
            ifont.FontHeightInPoints = 10;
            ifont.Boldweight = (short)FontBoldWeight.Normal;
            ifont.Color = HSSFColor.Grey40Percent.Index;
            cellStyle.SetFont(ifont);

            return cellStyle;
        }

        /// <summary>
        /// 设置默认页头样式
        /// </summary>
        /// <param name="workbook"></param>
        /// <param name="cell"></param>
        public static void SetDefaultHeaderCellStyle(IWorkbook workbook, ICell cell)
        {
            ICellStyle cellStyle = GetDefaultHeaderCellStyle(workbook);
            cell.CellStyle = cellStyle;
        }

        public static ICellStyle GetDefaultFooterCellStyle(IWorkbook workbook)
        {
            return null;
        }

        /// <summary>
        /// 设置默认页尾样式
        /// </summary>
        /// <param name="workbook"></param>
        /// <param name="cell"></param>
        public static void SetDefaultFooterCellStyle(IWorkbook workbook, ICell cell)
        {
            ICellStyle cellStyle = GetDefaultFooterCellStyle(workbook);
            cell.CellStyle = cellStyle;
        }
    }
}
