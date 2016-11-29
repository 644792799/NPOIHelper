using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace NPOIHelper.Contract
{
    public interface IPrint
    {
        /// <summary>
        /// 本地EXCEL文件存在的情况下打印
        /// </summary>
        /// <param name="localFilePath">本地EXCEL文件路径</param>
        /// <param name="strSheetName">SHEET名</param>
        /// <param name="callback">回调</param>
        void ExcelPrint(string localFilePath, string strSheetName, IPrintCallback callback);

        /// <summary>
        /// 从远端获取EXCEL文件打印
        /// </summary>
        /// <param name="remoteFilePath">远端文件路径（@"/127.0.0.1/print/example.xls"）</param>
        /// <param name="localPath">本地保存路径</param>
        /// <param name="strSheetName">SHEET名</param>
        /// <param name="callback">回调</param>
        void ExcelPrint(string remoteFilePath, string localPath, string strSheetName, IPrintCallback callback);
    }
}
