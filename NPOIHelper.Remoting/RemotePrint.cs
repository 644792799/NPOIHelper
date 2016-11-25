using NPOIHelper.Contract;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace NPOIHelper.Remoting
{
    public class RemotePrint : MarshalByRefObject, IPrint
    {

        #region IPrint 成员

        public void ExcelPrint(string strFilePath, string strSheetName, IPrintCallback callback)
        {
            Microsoft.Office.Interop.Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel.Workbook xlWorkbook;
            Microsoft.Office.Interop.Excel.Worksheet xlWorksheet;
            bool retry = false;
            int retryCount = 0;
            try
            {
                //string exeFolder = Path.GetDirectoryName(Assembly.GetEntryAssembly().Location);
                //strFilePath = Path.Combine(exeFolder, strFilePath);

                object oMissing = System.Reflection.Missing.Value;
                //strFilePath = Server.MapPath(strFilePath);
                if (!System.IO.File.Exists(strFilePath))
                {
                    throw new System.IO.FileNotFoundException();
                    //return;
                }
            
                xlApp.Visible = true;
                do
                {
                    try
                    {
                        xlWorkbook = xlApp.Workbooks.Add(strFilePath);
                        xlWorksheet = (Microsoft.Office.Interop.Excel.Worksheet)xlWorkbook.Worksheets[strSheetName];
                        xlWorksheet.PrintPreview(false);
                        xlWorkbook.Close(oMissing, oMissing, oMissing);
                        retry = false;
                    }
                    catch (Exception e)
                    {
                        retry = true;
                    }
                    finally
                    {
                        retryCount++;
                        if (retryCount > 5)
                        {
                            retry = false;
                            throw new System.Runtime.InteropServices.COMException();
                        }
                    }
                } while (retry);
                callback.PrintSuccess(true);
            }
            catch (Exception ex)
            {
                //throw ex;
                callback.PrintSuccess(false);
            }
            finally
            {
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
    }
}
