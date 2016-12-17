using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace NPOIHelper.NPOI.Excel
{
    public static class ExcelCellStyle
    {
        public static void Style(this ICell cell, string style)
        {
            Dictionary<string, string> dicproperties = style.ToDic();
            IWorkbook workbook = cell.Sheet.Workbook;
            ICellStyle cellstyle;
            if (cell.CellStyle == null)
            {
                cellstyle = workbook.CreateCellStyle();
            }
            else
            {
                cellstyle = workbook.CreateCellStyle();
                cellstyle.CloneStyleFrom(cell.CellStyle);
            }
            cell.CellStyle = cellstyle;
        }

        public static Dictionary<string, string> ToDic(this string style)
        {
            Dictionary<string, string> dicproperties = new Dictionary<string, string>();
            if (style.EndsWith(";"))
            {
                style = style.TrimEnd(new char[] { ';' });
            }
            string[] propertiesArr = style.Split(new char[] { ';' });
            foreach (var p in propertiesArr)
            {
                string[] kv = p.Split(new char[] { ':' });
                string key = kv[0].Trim();
                string value = kv[1].Trim();
                if (!dicproperties.ContainsKey(key) && Common.CommonConst.STANDARD_STYLE_PROPERTIES.Contains(key))
                {
                    dicproperties.Add(key, value);
                }
            }
            return dicproperties;
        }

        public static void AttachStyleToCellStyle(ICellStyle cellstyle, Dictionary<string, string> prop)
        {
            foreach (var p in prop)
            {
                cellstyle.AttachProperty(p.Key, p.Value);
            }
        }

        public static void AttachProperty(this ICellStyle cellstyle, string key, string value)
        {
            switch (key)
            {
                case "border-top":
                    cellstyle.BorderTop = (BorderStyle)Enum.ToObject(typeof(BorderStyle), value);
                    break;
                case "border-right":
                    cellstyle.BorderRight = (BorderStyle)Enum.ToObject(typeof(BorderStyle), value);
                    break;
                case "border-bottom":
                    cellstyle.BorderBottom = (BorderStyle)Enum.ToObject(typeof(BorderStyle), value);
                    break;
                case "border-left":
                    cellstyle.BorderLeft = (BorderStyle)Enum.ToObject(typeof(BorderStyle), value);
                    break;
                case "border-top-color":

                    break;
                case "border-right-color":

                    break;
                case "border-bottom-color":

                    break;
                case "border-left-color":

                    break;
                case "background-color":

                    break;
                case "border-diagonal":

                    break;
                case "border-diagonal-color":

                    break;
                case "border-diagonal-lines-tyle":

                    break;
                case "color":

                    break;
                case "data-format":

                    break;
                case "vertical-align":

                    break;
                case "text-align":

                    break;
            }
        }
    }
}
