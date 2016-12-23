using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace NPOIHelper.NPOI.Abstract
{
    public abstract class Title
    {
        const short DEFAULT_TITLE_HEIGHT = 30;
        public Title()
        {
            TitleHeight = DEFAULT_TITLE_HEIGHT;
        }

        public string DefaultTitleCellStyle { get; set; }

        public string TableTitle { get; set; }

        public short TitleHeight { get; set; }

        public bool IsNullOrWhiteSpace()
        {
            return string.IsNullOrWhiteSpace(TableTitle);
        }
    }
}
