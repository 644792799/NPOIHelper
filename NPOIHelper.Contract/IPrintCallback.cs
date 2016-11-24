using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace NPOIHelper.Contract
{
    public interface IPrintCallback
    {
        /// <summary>
        /// 回调方法，返回是否打印成功
        /// </summary>
        /// <returns></returns>
        void PrintSuccess(bool issuccess);
    }
}
