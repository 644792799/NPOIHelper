using NPOIHelper.Contract;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Remoting;
using System.Text;
using System.Threading.Tasks;

namespace NPOIHelper.Client
{
    class Program
    {
        static IPrint proxy = null;
        static string message = string.Empty;

        static void Main(string[] args)
        {
            //
            //RemotingConfiguration.Configure("NPOIHelper.Client.exe.config", false);
            RemotingConfiguration.Configure(AppDomain.CurrentDomain.SetupInformation.ConfigurationFile, false);
            proxy = (IPrint)Activator.GetObject(typeof(IPrint), "http://localhost:1234/Print/PrintURL");

            //message = Console.ReadLine();
            Print();
        }

        private static void Print()
        {
            proxy.ExcelPrint("d:\\aa\\test.xls", "疑似黑广播信号出现情况1", new PrintCallBackHandler());
            //proxy.SendMessage(message, new ChatRoomCallBackHandler());

            //message = Console.ReadLine();
            //if (message != "exit")
            //{
            //    SendMessage(message);
            //}
        }
    }
}
