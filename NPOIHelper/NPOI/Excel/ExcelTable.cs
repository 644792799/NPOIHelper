using NPOIHelper.NPOI.Abstract;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace NPOIHelper.NPOI.Excel
{
    public class ExcelTable : Table
    {
        public ExcelTable()
            : base()
        {
            //Rows = new List<Row>();
        }

        public override Row CreateRow(bool isRowHeader = false)
        {
            ExcelRow row = new ExcelRow();
            row.IsRowHead = isRowHeader;
            //row.IsColHead = isColHeader;
            return row;
        }

        public override void AddRow(Row row)
        {
            Rows.Add(row);
        }

        public override void AddRow(int rowIndex)
        {
            throw new NotImplementedException();
        }

        public override void DeleteRow(int rowIndex)
        {
            throw new NotImplementedException();
        }

        public override void DeleteRows(int fromRowIndex, int toRowIndex)
        {
            throw new NotImplementedException();
        }

        public void ExportToExcel()
        {

        }
    }
}
