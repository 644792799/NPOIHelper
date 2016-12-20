using NPOI.HSSF.Record;
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
    /// <summary>
    ///  ┏━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━┓
    ///  ┃                        TITLE                            ┃
    ///  ┣━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━┫
    ///  ┃                        HEADER                           ┃
    ///  ┣━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━┫
    ///  ┃    ┃    ┃    ┃    ┃TABLE HEADER┃    ┃    ┃    ┃    ┃    ┃
    ///  ┣━━━━╋━━━━╋━━━━╋━━━━╋━━━━━━━━━━━━╋━━━━╋━━━━╋━━━━╋━━━━╋━━━━┫
    ///  ┃    ┃    ┃    ┃    ┃TABLE   BODY┃    ┃    ┃    ┃    ┃    ┃
    ///  ┣━━━━╋━━━━╋━━━━╋━━━━╋━━━━━━━━━━━━╋━━━━╋━━━━╋━━━━╋━━━━╋━━━━┫
    ///  ┃    ┃    ┃    ┃    ┃TABLE   BODY┃    ┃    ┃    ┃    ┃    ┃
    ///  ┣━━━━╋━━━━╋━━━━╋━━━━╋━━━━━━━━━━━━╋━━━━╋━━━━╋━━━━╋━━━━╋━━━━┫
    ///  ┃    ┃    ┃    ┃    ┃TABLE FOOTER┃    ┃    ┃    ┃    ┃    ┃
    ///  ┣━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━┫
    ///  ┃                        FOOTER                           ┃
    ///  ┗━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━┛
    /// </summary>
    public static class ExcelCellStyle
    {
        #region 样式扩展
        const string DEFAULT_CELLSTYLE = "border:thin;text-align:center;vertical-align:center;font-name:宋体;font-size:9;font-weight:normal;font-color:GREY_50_PERCENT;";
        const string DEFAULT_TITLE_CELLSTYLE = "border:none;text-align:center;vertical-align:center;wraptext:true;font-name:方正姚体;font-size:20;font-weight:bold;font-color:GREY_50_PERCENT;";
        const string DEFAULT_HEADER_CELLSTYLE = "border:none;text-align:left;vertical-align:center;wraptext:true;font-name:宋体;font-size:10;font-weight:normal;font-color:GREY_40_PERCENT;";
        const string DEFAULT_TABLE_HEADER_CELLSTYLE = "border:thin;text-align:center;vertical-align:center;pattern:SolidForeground;color:Grey25Percent;wraptext:true;font-name:方正姚体;font-size:12;font-weight:bold;font-color:black;";
        const string DEFAULT_TABLE_BODY_CELLSTYLE = "border:thin;text-align:center;vertical-align:center;font-name:宋体;font-size:9;font-weight:normal;font-color:GREY_50_PERCENT;";
        const string DEFAULT_TABLE_FOOTER_CELLSTYLE = "";
        const string DEFAULT_FOOTER_CELLSTYLE = "border:none;text-align:center;vertical-align:center;font-name:宋体;font-size:9;font-weight:normal;font-color:GREY_50_PERCENT;";
        
        private static ICellStyle defaultHeaderCellStyle;
        /// <summary>
        /// 获取默认页头样式
        /// </summary>
        /// <param name="workbook"></param>
        /// <returns></returns>
        public static ICellStyle GetDefaultHeaderCellStyle(this IWorkbook workbook)
        {
            if (defaultHeaderCellStyle == null)
            {
                defaultHeaderCellStyle = workbook.CreateCellStyle();
                defaultHeaderCellStyle.AttachStyleToCellStyle(workbook, DEFAULT_HEADER_CELLSTYLE);
            }
            return defaultHeaderCellStyle;
        }
        /// <summary>
        /// 属性font-前缀的属性转IFont
        /// </summary>
        /// <param name="workbook"></param>
        /// <param name="fontDic"></param>
        /// <returns></returns>
        public static IFont GetFont(this IWorkbook workbook, Dictionary<string, string> fontDic)
        {
            IFont font = workbook.CreateFont();
            fontDic.FontWeight(font);
            fontDic.FontColor(font);
            fontDic.FontItalic(font);
            fontDic.FontName(font);
            fontDic.FontSize(font);
            fontDic.FontStrikeout(font);
            fontDic.FontOffset(font);
            fontDic.FontUnderline(font);
            return font;
        }
        /// <summary>
        /// 通过CSS形式的文本添加单元格样式
        /// </summary>
        /// <param name="cell"></param>
        /// <param name="style"></param>
        public static void Style(this ICell cell, string style)
        {
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
            cellstyle.AttachStyleToCellStyle(workbook, style);
            cell.CellStyle = cellstyle;
        }
        /// <summary>
        /// CSS格式属性赋值给单元格样式
        /// </summary>
        /// <param name="cellstyle"></param>
        /// <param name="workbook"></param>
        /// <param name="style"></param>
        public static void AttachStyleToCellStyle(this ICellStyle cellstyle, IWorkbook workbook, string style)
        {
            Dictionary<string, string> prop = style.ToDic();
            var fontStyles = prop.Where(w => w.Key.StartsWith("font-")).ToArray();
            var fontDic = new Dictionary<string, string>();
            foreach (var kv in fontStyles)
            {
                fontDic.Add(kv.Key, kv.Value);
            }
            var font = workbook.GetFont(fontDic);
            cellstyle.SetFont(font);
            var propDic = prop.Except(fontDic);
            foreach (var kvp in propDic)
            {
                cellstyle.AttachProperty(workbook, kvp.Key, kvp.Value);
            }
        }
        /// <summary>
        /// 附加cellstyle样式
        /// </summary>
        /// <param name="cellstyle"></param>
        /// <param name="workbook"></param>
        /// <param name="key"></param>
        /// <param name="value"></param>
        public static void AttachProperty(this ICellStyle cellstyle, IWorkbook workbook, string key, string value)
        {
            switch (key)
            {
                case "border":
                    BorderStyle bs = value.ConvertToBorderStyle();
                    cellstyle.BorderTop = bs;
                    cellstyle.BorderRight = bs;
                    cellstyle.BorderBottom = bs;
                    cellstyle.BorderLeft = bs;
                    break;
                case "border-top":
                    cellstyle.BorderTop = value.ConvertToBorderStyle();//(BorderStyle)Enum.ToObject(typeof(BorderStyle), value);
                    break;
                case "border-right":
                    cellstyle.BorderRight = value.ConvertToBorderStyle();//(BorderStyle)Enum.ToObject(typeof(BorderStyle), value);
                    break;
                case "border-bottom":
                    cellstyle.BorderBottom = value.ConvertToBorderStyle();//(BorderStyle)Enum.ToObject(typeof(BorderStyle), value);
                    break;
                case "border-left":
                    cellstyle.BorderLeft = value.ConvertToBorderStyle();//(BorderStyle)Enum.ToObject(typeof(BorderStyle), value);
                    break;
                case "border-color":
                    short xlcolor = GetXLColor((HSSFWorkbook)workbook, value);
                    cellstyle.TopBorderColor = xlcolor;
                    cellstyle.RightBorderColor = xlcolor;
                    cellstyle.BottomBorderColor = xlcolor;
                    cellstyle.LeftBorderColor = xlcolor;
                    break;
                case "border-top-color":
                    cellstyle.TopBorderColor = GetXLColor((HSSFWorkbook)workbook, value);
                    break;
                case "border-right-color":
                    cellstyle.RightBorderColor = GetXLColor((HSSFWorkbook)workbook, value);
                    break;
                case "border-bottom-color":
                    cellstyle.BottomBorderColor = GetXLColor((HSSFWorkbook)workbook, value);
                    break;
                case "border-left-color":
                    cellstyle.LeftBorderColor = GetXLColor((HSSFWorkbook)workbook, value);
                    break;
                case "background-color":
                    cellstyle.FillBackgroundColor = GetXLColor((HSSFWorkbook)workbook, value);
                    break;
                case "border-diagonal":

                    break;
                case "border-diagonal-color":
                    cellstyle.BorderDiagonalColor = GetXLColor((HSSFWorkbook)workbook, value);
                    break;
                case "border-diagonal-lines-tyle":

                    break;
                case "color":
                    cellstyle.FillForegroundColor = GetXLColor((HSSFWorkbook)workbook, value);
                    break;
                case "pattern":
                    cellstyle.FillPattern = value.ConvertToFillPattern();
                    break;
                case "data-format":
                    
                    break;
                case "vertical-align":
                    cellstyle.VerticalAlignment = value.ConvertToVerticalAlignment();
                    break;
                case "text-align":
                    cellstyle.Alignment = value.ConvertToHorizontalAlignment();
                    break;
                case "wraptext":
                    cellstyle.WrapText = IsWrapText(value);
                    break;
            }
        }
        #endregion

        #region 工具方法
        /// <summary>
        /// 颜色字符串转short类型
        /// </summary>
        /// <param name="workbook"></param>
        /// <param name="htmlcolor"></param>
        /// <returns></returns>
        private static short GetXLColor(HSSFWorkbook workbook, string htmlcolor)
        {
            HSSFPalette XlPalette = workbook.GetCustomPalette();
            if (htmlcolor.StartsWith("#"))
            {
                System.Drawing.Color color = System.Drawing.ColorTranslator.FromHtml(htmlcolor);
                HSSFColor hssfcolor = XlPalette.FindColor(color.R, color.G, color.B);
                if (hssfcolor == null)
                {
                    hssfcolor = XlPalette.AddColor(color.R, color.G, color.B);
                }
                return hssfcolor.Indexed;
            }
            else
            {
                return htmlcolor.ToUpper().ConvertToColor();
            }
        }
        #endregion

        #region 字符串扩展
        /// <summary>
        /// CSS格式的字符串转DICTIONARY
        /// </summary>
        /// <param name="style"></param>
        /// <returns></returns>
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
        public static short ConvertToColor(this string v)
        {
            switch (v)
            {
                case "AQUA":
                    return 49;

                case "AUTOMATIC":
                    return 64;

                case "BLACK":
                    return 8;

                case "BLUE":
                    return 12;

                case "BLUE_GREY":
                case "BLUEGREY":
                    return 54;

                case "BRIGHT_GREEN":
                case "BRIGHTGREEN":
                    return 11;

                case "BROWN":
                    return 60;

                case "CORAL":
                    return 29;

                case "CORNFLOWER_BLUE":
                case "CORNFLOWERBLUE":
                    return 24;

                case "DARK_BLUE":
                case "DARKBLUE":
                    return 18;

                case "DARK_GREEN":
                case "DARKGREEN":
                    return 58;

                case "DARK_RED":
                case "DARKRED":
                    return 16;

                case "DARK_TEAL":
                case "DARKTEAL":
                    return 56;

                case "DARK_YELLOW":
                case "DARKYELLOW":
                    return 19;

                case "GOLD":
                    return 51;

                case "GREEN":
                    return 17;

                case "GREY_25_PERCENT":
                case "GREY25PERCENT":
                    return 22;

                case "GREY_40_PERCENT":
                case "GREY40PERCENT":
                    return 55;

                case "GREY_50_PERCENT":
                case "GREY50PERCENT":
                    return 23;

                case "GREY_80_PERCENT":
                case "GREY80PERCENT":
                    return 63;

                case "INDIGO":
                    return 62;

                case "LAVENDER":
                    return 46;

                case "LEMON_CHIFFON":
                case "LEMONCHIFFON":
                    return 26;

                case "LIGHT_BLUE":
                case "LIGHTBLUE":
                    return 48;

                case "LIGHT_CORNFLOWERBLUE":
                case "LIGHTCORNFLOWERBLUE":
                    return 31;

                case "LIGHT_GREEN":
                case "LIGHTGREEN":
                    return 42;

                case "LIGHT_ORANGE":
                case "LIGHTORANGE":
                    return 52;

                case "LIGHT_TURQUOISE":
                case "LIGHTTURQUOISE":
                    return 41;

                case "LIGHT_YELLOW":
                case "LIGHTYELLOW":
                    return 43;

                case "LIME":
                    return 50;

                case "MAROON":
                    return 25;

                case "OLIVE_GREEN":
                case "OLIVEGREEN":
                    return 59;

                case "ORANGE":
                    return 53;

                case "ORCHID":
                    return 28;

                case "PALE_BLUE":
                case "PALEBLUE":
                    return 44;

                case "PINK":
                    return 14;

                case "PLUM":
                    return 61;

                case "RED":
                    return 10;

                case "ROSE":
                    return 45;

                case "ROYAL_BLUE":
                case "ROYALBLUE":
                    return 30;

                case "SEA_GREEN":
                case "SEAGREEN":
                    return 57;

                case "SKY_BLUE":
                case "SKYBLUE":
                    return 40;

                case "TAN":
                    return 47;

                case "TEAL":
                    return 21;

                case "TURQUOISE":
                    return 15;

                case "VIOLET":
                    return 20;

                case "WHITE":
                    return 9;

                case "YELLOW":
                    return 13;

                default:
                    return 32767;
            }
        }

        public static HorizontalAlignment ConvertToHorizontalAlignment(this string v)
        {
            switch (v.ToUpper())
            {
                case "LEFT":
                    return HorizontalAlignment.Left;

                case "CENTER":
                    return HorizontalAlignment.Center;

                case "CENTERSELECTION":
                case "CENTER_SELECTION":
                    return HorizontalAlignment.CenterSelection;

                case "RIGHT":
                    return HorizontalAlignment.Right;

                case "DISTRIBUTED":
                    return HorizontalAlignment.Distributed;

                case "FILL":
                    return HorizontalAlignment.Fill;

                case "JUSTIFY":
                    return HorizontalAlignment.Justify;

                default:
                    return HorizontalAlignment.General;
            }
        }

        public static VerticalAlignment ConvertToVerticalAlignment(this string v)
        {
            switch (v.ToUpper())
            {
                case "TOP":
                    return VerticalAlignment.Top;

                case "CENTER":
                    return VerticalAlignment.Center;

                case "BOTTOM":
                    return VerticalAlignment.Bottom;

                case "DISTRIBUTED":
                    return VerticalAlignment.Distributed;

                default:
                    return VerticalAlignment.Justify;
            }
        }

        public static BorderStyle ConvertToBorderStyle(this string v)
        {
            int a = 0;
            if (int.TryParse(v, out a))
            {
                if (a < 0 || a > 13)
                {
                    return BorderStyle.None;
                }
                else
                {
                    return (BorderStyle)Enum.ToObject(typeof(BorderStyle), v);
                }
            }
            else
            {
                switch (v.ToUpper())
                {
                    case "THIN":
                        return BorderStyle.Thin;

                    case "MEDIUM":
                        return BorderStyle.Medium;

                    case "DASHED":
                        return BorderStyle.Dashed;

                    case "HAIR":
                        return BorderStyle.Hair;

                    case "THICK":
                        return BorderStyle.Thick;

                    case "DOUBLE":
                        return BorderStyle.Double;

                    case "DOTTED":
                        return BorderStyle.Dotted;

                    case "MEDIUMDASHED":
                    case "MEDIUM_DASHED":
                        return BorderStyle.MediumDashed;

                    case "DASHDOT":
                    case "DASH_DOT":
                        return BorderStyle.DashDot;

                    case "MEDIUMDASHDOT":
                    case "MEDIUM_DASH_DOT":
                        return BorderStyle.MediumDashDot;

                    case "DASHDOTDOT":
                    case "DASH_DOT_DOT":
                        return BorderStyle.DashDotDot;

                    case "MEDIUMDASHDOTDOT":
                    case "MEDIUM_DASH_DOT_DOT":
                        return BorderStyle.MediumDashDotDot;

                    case "SLANTEDDASHDOT":
                    case "SLANTED_DASH_DOT":
                        return BorderStyle.SlantedDashDot;

                    default:
                        return BorderStyle.None;
                }
            }
        }

        public static FillPattern ConvertToFillPattern(this string v)
        {
            int a = 0;
            if (int.TryParse(v, out a))
            {
                if (a < 0 || a > 18)
                {
                    return FillPattern.NoFill;
                }
                else
                {
                    return (FillPattern)Enum.ToObject(typeof(FillPattern), a);
                }
            }
            else
            {
                switch (v.ToUpper())
                {
                    case "NOFILL"://0,
                        return FillPattern.NoFill;
                    case "SOLIDFOREGROUND"://1,
                        return FillPattern.SolidForeground;
                    case "FINEDOTS"://2,
                        return FillPattern.FineDots;
                    case "ALTBARS"://3,
                        return FillPattern.AltBars;
                    case "SPARSEDOTS"://4,
                        return FillPattern.SparseDots;
                    case "THICKHORIZONTALBANDS"://5,
                        return FillPattern.ThickHorizontalBands;
                    case "THICKVERTICALBANDS"://6,
                        return FillPattern.ThickVerticalBands;
                    case "THICKBACKWARDDIAGONALS"://7,
                        return FillPattern.ThickBackwardDiagonals;
                    case "THICKFORWARDDIAGONALS"://8,
                        return FillPattern.ThickForwardDiagonals;
                    case "BIGSPOTS"://9,
                        return FillPattern.BigSpots;
                    case "BRICKS"://10,
                        return FillPattern.Bricks;
                    case "THINHORIZONTALBANDS"://11,
                        return FillPattern.ThinHorizontalBands;
                    case "THINVERTICALBANDS"://12,
                        return FillPattern.ThinVerticalBands;
                    case "THINBACKWARDDIAGONALS"://13,
                        return FillPattern.ThinBackwardDiagonals;
                    case "THINFORWARDDIAGONALS"://14,
                        return FillPattern.ThinForwardDiagonals;
                    case "SQUARES"://15,
                        return FillPattern.Squares;
                    case "DIAMONDS"://16,
                        return FillPattern.Diamonds;
                    case "LESSDOTS"://17,
                        return FillPattern.LessDots;
                    case "LEASTDOTS"://18,
                        return FillPattern.LeastDots;
                    default:
                        return FillPattern.NoFill;
                        break;
                }
            }
        }
        public static bool IsWrapText(string v)
        {
            return v.ToUpper() == "TRUE";
        }
        #endregion

        #region 字典扩展
        public static string GetDicValue(Dictionary<string, string> dic, string key)
        {
            if (dic.ContainsKey(key))
            {
                return dic[key];
            }
            else
            {
                return null;
            }
        }

        public static void FontWeight(this Dictionary<string, string> fontdic, IFont font)
        {
            string fontweight = GetDicValue(fontdic, "font-weight");
            switch (fontweight)
            {
                case "NORMAL":
                    font.Boldweight = 400;
                    break;
                case "BOLD":
                    font.Boldweight = 700;
                    break;
                default:
                    font.Boldweight = short.Parse(fontweight);
                    break;
            }
        }

        public static void FontName(this Dictionary<string, string> fontdic, IFont font)
        {
            string fontname = GetDicValue(fontdic, "font-name");
            if (fontname != null)
            {
                font.FontName = fontname;
            }
        }

        public static void FontSize(this Dictionary<string, string> fontdic, IFont font)
        {
            string fontsize = GetDicValue(fontdic, "font-size");
            if (fontsize != null)
            {
                font.FontHeightInPoints = short.Parse(fontsize);
            }
        }

        public static void FontColor(this Dictionary<string, string> fontdic, IFont font)
        {
            string fontcolor = GetDicValue(fontdic, "font-color");
            if (fontcolor != null)
            {
                font.Color = fontcolor.ConvertToColor();
            }
        }

        public static void FontItalic(this Dictionary<string, string> fontdic, IFont font)
        {
            string fontitalic = GetDicValue(fontdic, "font-italic");
            if (fontitalic != null)
            {
                font.IsItalic = fontitalic.ToUpper() == "TRUE";
            }
        }

        public static void FontStrikeout(this Dictionary<string, string> fontdic, IFont font)
        {
            string fontstrikeout = GetDicValue(fontdic, "font-strikeout");
            if (fontstrikeout != null)
            {
                font.IsItalic = fontstrikeout.ToUpper() == "TRUE";
            }
        }

        public static void FontUnderline(this Dictionary<string, string> fontdic, IFont font)
        {
            string fontunderline = GetDicValue(fontdic, "font-underline");
            if (fontunderline != null)
            {
                switch (fontunderline.ToUpper())
                {
                    case "SINGLE":
                        font.Underline = (FontUnderlineType)Enum.ToObject(typeof(FontUnderlineType), 1);
                        break;
                    case "DOUBLE":
                        font.Underline = (FontUnderlineType)Enum.ToObject(typeof(FontUnderlineType), 2);
                        break;
                    case "SINGLEACCOUNTING":
                    case "SINGLE_ACCOUNTING":
                        font.Underline = (FontUnderlineType)Enum.ToObject(typeof(FontUnderlineType), 33);
                        break;
                    case "DOUBLEACCOUNTING":
                    case "DOUBLE_ACCOUNTING":
                        font.Underline = (FontUnderlineType)Enum.ToObject(typeof(FontUnderlineType), 34);
                        break;
                    case "NONE":
                        font.Underline = (FontUnderlineType)Enum.ToObject(typeof(FontUnderlineType), 0);
                        break;
                    default:
                        break;
                }
            }
        }

        public static void FontOffset(this Dictionary<string, string> fontdic, IFont font)
        {
            string fontoffset = GetDicValue(fontdic, "font-offset");
            if (fontoffset != null)
            {
                switch (fontoffset.ToUpper())
                {
                    case "SUPER":
                        font.TypeOffset = (FontSuperScript)Enum.ToObject(typeof(FontSuperScript), 1);
                        break;
                    case "SUB":
                        font.TypeOffset = (FontSuperScript)Enum.ToObject(typeof(FontSuperScript), 2);
                        break;
                    case "NONE":
                        font.TypeOffset = (FontSuperScript)Enum.ToObject(typeof(FontSuperScript), 0);
                        break;
                    default:
                        break;
                }
            }
        }
        #endregion
    }
}
