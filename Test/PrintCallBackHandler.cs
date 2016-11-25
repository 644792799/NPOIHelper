using NPOIHelper.Contract;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Test
{
    public class PrintCallBackHandler : MarshalByRefObject, IPrintCallback
    {
        public void PrintSuccess(bool issuccess)
        {
            Console.WriteLine(issuccess);
            //throw new NotImplementedException();
        }
    }
}
