using NPOIHelper.NPOI.Abstract;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace NPOIHelper.NPOI.Excel
{
    class ExcelRow : Row
    {
        public override int CreateCell()
        {
            
            throw new NotImplementedException();
        }

        public override void AddCell()
        {
            throw new NotImplementedException();
        }

        public override void AddCell(int columnIndex)
        {
            throw new NotImplementedException();
        }

        public override void DeleteCell(int columnIndex)
        {
            throw new NotImplementedException();
        }

        public override void DeleteCells(int fromColumnIndex, int toColumnIndex)
        {
            throw new NotImplementedException();
        }
    }
}
