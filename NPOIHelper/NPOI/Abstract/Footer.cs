﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace NPOIHelper.NPOI.Abstract
{
    public abstract class Footer
    {
        public Footer()
        {
            Rows = new List<Row>();
        }
        public IList<Row> Rows { get; set; }
    }
}
