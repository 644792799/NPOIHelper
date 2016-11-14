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
        public static void SetCellStyle(ICell cell, ICellStyle cellstyle)
        {
            cell.CellStyle = cellstyle;
        }

        public static ICellStyle GetDefaultCellStyle(IWorkbook wb)
        {
            ICellStyle cellStyle = wb.CreateCellStyle();
            IFont ifont = wb.CreateFont();
            ifont.FontName = "宋体";
            ifont.FontHeightInPoints = 9;
            cellStyle.BorderLeft = BorderStyle.Thin;
            cellStyle.BorderRight = BorderStyle.Thin;
            cellStyle.BorderTop = BorderStyle.Thin;
            cellStyle.BorderBottom = BorderStyle.Thin;
            cellStyle.LeftBorderColor = HSSFColor.Black.Index;
            return cellStyle;
        }

        public static void SetDefaultCellStyle(IWorkbook workbook, ICell cell)
        {
            ICellStyle cellStyle = GetDefaultCellStyle(workbook);
            cell.CellStyle = cellStyle;
        }

        public static void SetDefaultTitleCellStyle(IWorkbook workbook, ICell cell)
        {
            ICellStyle cellStyle = GetDefaultCellStyle(workbook);
            IFont ifont = workbook.CreateFont();
            cellStyle.VerticalAlignment = VerticalAlignment.Center;
            ifont.FontHeightInPoints = 12;
            ifont.Boldweight = (short)FontBoldWeight.Bold;
            cellStyle.SetFont(ifont);

            cell.CellStyle = cellStyle;
        }

        public static void SetDefaultTableHeaderCellStyle(IWorkbook workbook, ICell cell)
        {
            ICellStyle cellStyle = GetDefaultCellStyle(workbook);
            IFont ifont = workbook.CreateFont();
            //背景
            cellStyle.FillPattern = FillPattern.SolidForeground;
            cellStyle.FillForegroundColor = HSSFColor.LemonChiffon.Index;
            //水平对齐 
            cellStyle.Alignment = HorizontalAlignment.Center;
            //垂直对齐
            cellStyle.VerticalAlignment = VerticalAlignment.Center;
            //边框  
            cellStyle.BorderLeft = BorderStyle.None;
            cellStyle.BorderRight = BorderStyle.None;
            cellStyle.BorderTop = BorderStyle.None;
            cellStyle.BorderBottom = BorderStyle.None;
            //字体样式
            ifont.FontHeightInPoints = 16;
            ifont.Boldweight = (short)FontBoldWeight.Bold;
            ifont.Color = HSSFColor.Green.Index;
            cellStyle.SetFont(ifont);

            cell.CellStyle = cellStyle;
        }

        public static void SetDefaultTableBodyCellStyle(IWorkbook workbook, ICell cell)
        {

        }

        public static void SetDefaultTableFooterCellStyle(IWorkbook workbook, ICell cell)
        {

        }

        public static void SetDefaultHeaderCellStyle(IWorkbook workbook, ICell cell)
        {

        }

        public static void SetDefaultFooterCellStyle(IWorkbook workbook, ICell cell)
        {

        }
    }
}
