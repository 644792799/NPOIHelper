using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace NPOIHelper.Contract
{
    public interface IPrint
    {
        void ExcelPrint(string strFilePath, string strSheetName, IPrintCallback callback);
    }
}
