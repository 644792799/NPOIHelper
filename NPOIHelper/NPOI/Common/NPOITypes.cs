﻿using NPOI.HSSF.Util;
using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace NPOIHelper.NPOI.Common
{
    public enum CellTypes
    {
        String = CellType.String,
        Numeric = CellType.Numeric,
        Image = 6
    }
    public enum TextAlign
    {
        Center,
        Left,
        Right
    }
    public enum DefaultCellStyle
    {
        Title,
        Header,
        Footer,
        TableHeader,
        TableBody,
        TableFooter
    }
    public enum FontColors
    {
        Aqua = HSSFColor.Aqua.Index,
        Automatic = HSSFColor.Automatic.Index,
        Black = HSSFColor.Black.Index,
        Blue = HSSFColor.Blue.Index,
        BlueGrey = HSSFColor.BlueGrey.Index,
        BrightGreen = HSSFColor.BrightGreen.Index,
        Brown = HSSFColor.Brown.Index,
        Coral = HSSFColor.Coral.Index,
        CornflowerBlue = HSSFColor.CornflowerBlue.Index,
        DarkBlue = HSSFColor.DarkBlue.Index,
        DarkGreen = HSSFColor.DarkGreen.Index,
        DarkRed = HSSFColor.DarkRed.Index,
        DarkTeal = HSSFColor.DarkTeal.Index,
        DarkYellow = HSSFColor.DarkYellow.Index,
        Gold = HSSFColor.Gold.Index,
        Green = HSSFColor.Green.Index,
        Indigo = HSSFColor.Indigo.Index,
        Lavender = HSSFColor.Lavender.Index,
        LemonChiffon = HSSFColor.LemonChiffon.Index,
        LightBlue = HSSFColor.LightBlue.Index,
        LightCornflowerBlue = HSSFColor.LightCornflowerBlue.Index,
        LightGreen = HSSFColor.LightGreen.Index,
        LightOrange = HSSFColor.LightOrange.Index,
        LightTurquoise = HSSFColor.LightTurquoise.Index,
        LightYellow = HSSFColor.LightYellow.Index,
        Lime = HSSFColor.Lime.Index,
        Maroon = HSSFColor.Maroon.Index,
        OliveGreen = HSSFColor.OliveGreen.Index,
        Orange = HSSFColor.Orange.Index,
        Orchid = HSSFColor.Orchid.Index,
        PaleBlue = HSSFColor.PaleBlue.Index,
        Pink = HSSFColor.Pink.Index,
        Plum = HSSFColor.Plum.Index,
        Red = HSSFColor.Red.Index,
        Rose = HSSFColor.Rose.Index,
        RoyalBlue = HSSFColor.RoyalBlue.Index,
        SeaGreen = HSSFColor.SeaGreen.Index,
        SkyBlue = HSSFColor.SkyBlue.Index,
        Tan = HSSFColor.Tan.Index,
        Teal = HSSFColor.Teal.Index,
        Turquoise = HSSFColor.Turquoise.Index,
        Violet = HSSFColor.Violet.Index,
        White = HSSFColor.White.Index,
        Yellow = HSSFColor.Yellow.Index,
    }
}
