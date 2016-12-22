using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace NPOIHelper.NPOI.Abstract
{
    public abstract class Title
    {
        public Title()
        {
            
        }

        private string defaultTitleCellStyle;
        public string DefaultTitleCellStyle
        {
            get { return defaultTitleCellStyle; }
            set { defaultTitleCellStyle = value; }
        }

        public string TableTitle { get; set; }

        private short titleHeight = 30;
        public short TitleHeight
        {
            get { return titleHeight; }
            set { this.titleHeight = value; }
        }

        public bool IsNullOrWhiteSpace()
        {
            return string.IsNullOrWhiteSpace(TableTitle);
        }
    }
}
