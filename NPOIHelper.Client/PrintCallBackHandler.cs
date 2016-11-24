using NPOIHelper.Contract;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace NPOIHelper.Client
{
    public class PrintCallBackHandler : MarshalByRefObject, IPrintCallback
    {

        #region IPrintCallback 成员

        public void PrintSuccess(bool issuccess)
        {
            Console.WriteLine(issuccess);
            //throw new NotImplementedException();
        }

        #endregion
    }
}
