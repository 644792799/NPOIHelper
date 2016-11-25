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
            proxy = (IPrint)Activator.GetObject(typeof(IPrint), "http://172.39.8.173:1234/Print/PrintURL");

            //message = Console.ReadLine();
            Print();
        }

        private static void Print()
        {
            try
            {
                proxy.ExcelPrint("d:\\Debug\\疑似黑广播信号出现情况报表.xls", "疑似黑广播信号出现情况(成都站)", new PrintCallBackHandler());
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                Console.Read();
            }
            //proxy.SendMessage(message, new ChatRoomCallBackHandler());

            //message = Console.ReadLine();
            //if (message != "exit")
            //{
            //    SendMessage(message);
            //}
        }
    }
}
