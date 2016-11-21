using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace NPOIHelper.NPOI.Abstract
{
    public abstract class Table
    {
        const int A4_WIDTH = 794;
        const int A4_HEIGHT = 1123;
        /// <summary>
        /// 是否横向打印
        /// </summary>
        public bool Landscape { get; set; }
        public Header Header { get; set; }
        public string Title { get; set; }
        private short titleHeight = 30;
        public short TitleHeight {
            get { return titleHeight; }
            set { this.titleHeight = value; } 
        }
        public IList<Row> Rows { get; set; }
        public int ColumnCount { get; set; }
        public int[] columnWidths;
        public int[] ColumnWidths
        {
            get
            {
                return columnWidths;
            }
            set
            {
                //if (value == null)
                //{
                //    this.columnWidths = value;
                //    return;
                //}
                //List<int> list = value.ToList();
                //if (Landscape)
                //{
                //    for (int i = 0; i < value.Length; i++)
                //    {
                //        value[i] = value[i] * A4_HEIGHT / 100;
                //    }
                //}
                //else
                //{
                //    for (int i = 0; i < value.Length; i++)
                //    {
                //        value[i] = value[i] * A4_WIDTH / 100;
                //    }
                //}
                this.columnWidths = value;
            }
        }
        //public IList<string> Columns { get; set; }
        public Footer Footer { get; set; }

        public Table()
        {
            Rows = new List<Row>();
            Landscape = false;
        }

        abstract public Row CreateRow(bool isHeader);
        abstract public void AddRow(Row row);
        abstract public void AddRow(int rowIndex);
        abstract public void DeleteRow(int rowIndex);
        abstract public void DeleteRows(int fromRowIndex, int toRowIndex);

    }
}
