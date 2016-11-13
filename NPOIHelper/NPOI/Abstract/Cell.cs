using NPOIHelper.NPOI.Common;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace NPOIHelper.NPOI.Abstract
{
    public abstract class Cell
    {
        public dynamic Value { get; set; }
        public CellTypes CellType { get; set; }
        public short? FontColor { get; set; }

        public Cell(dynamic value) {
            Value = value;
            CellType = CellTypes.String;
        }
        public Cell(dynamic value,CellTypes cellType) {
            Value = value;
            CellType = cellType;
        }
        public Cell(dynamic value,FontColors fontColor) {
            Value = value;
            CellType = CellTypes.String;
            FontColor = (short)fontColor;
        }
        public Cell(dynamic value,short fontColor) {
            Value = value;
            CellType = CellTypes.String;
            FontColor = fontColor;
        }
        public Cell(dynamic value,CellTypes cellType,FontColors fontColor) {
            Value = value;
            CellType = cellType;
            FontColor = (short)fontColor;
        }
        public Cell(dynamic value, CellTypes cellType, short fontColor)
        {
            Value = value;
            CellType = cellType;
            FontColor = fontColor;
        }
    }
}
