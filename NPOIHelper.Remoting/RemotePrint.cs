using NPOIHelper.Contract;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace NPOIHelper.Remoting
{
    public class RemotePrint : MarshalByRefObject, IPrint
    {

        #region IPrint 成员

        public void ExcelPrint(string localFilePath, string strSheetName, IPrintCallback callback)
        {
            Microsoft.Office.Interop.Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel.Workbook xlWorkbook = null;
            Microsoft.Office.Interop.Excel.Worksheet xlWorksheet;
            object oMissing = System.Reflection.Missing.Value;
            //bool retry = false;
            //int retryCount = 0;
            try
            {
                if (!System.IO.File.Exists(localFilePath))
                {
                    throw new System.IO.FileNotFoundException();
                }
            
                xlApp.Visible = true;
                //do
                //{
                //    try
                //    {
                xlWorkbook = xlApp.Workbooks.Add(localFilePath);
                xlWorksheet = (Microsoft.Office.Interop.Excel.Worksheet)xlWorkbook.Worksheets[strSheetName];
                xlWorksheet.PrintPreview(false);
                //xlWorkbook.Close(oMissing, oMissing, oMissing);
                //retry = false;
                //    }
                //    catch (Exception e)
                //    {
                //        retry = true;
                //    }
                //    finally
                //    {
                //        retryCount++;
                //        if (retryCount > 5)
                //        {
                //            retry = false;
                //            throw new System.Runtime.InteropServices.COMException();
                //        }
                //    }
                //} while (retry);
                callback.PrintSuccess(true);
            }
            catch (Exception ex)
            {
                //throw ex;
                callback.PrintSuccess(false);
            }
            finally
            {
                if (xlWorkbook != null)
                {
                    xlWorkbook.Close(oMissing, oMissing, oMissing);
                }
                if (xlApp != null)
                {
                    xlApp.Quit();
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(xlApp);
                    xlApp = null;
                }
                GC.Collect();
            }
        }

        #endregion


        public void ExcelPrint(string remoteFilePath, string localPath, string strSheetName, IPrintCallback callback)
        {
            string localFilePath = SaveRemoteFileToLocal(remoteFilePath, localPath);
            ExcelPrint(localFilePath, strSheetName, callback);
        }

        /// <summary>
        /// 保存远程文件到本地
        /// </summary>
        /// <param name="remoteFilePath"></param>
        /// <param name="localPath"></param>
        /// <returns>localPath</returns>
        private string SaveRemoteFileToLocal(string remoteFilePath, string localPath)
        {
            localPath = localPath.EndsWith("/") || localPath.EndsWith("\\") ? localPath : localPath + "/";
            string localFilePath = localPath + System.IO.Path.GetFileName(remoteFilePath);
            if (!Directory.Exists(localPath))
            {
                Directory.CreateDirectory(localPath);
            }
            WebClient client = new WebClient();
            client.DownloadFile(@remoteFilePath, localFilePath);

            return localFilePath;
        }
    }
}
