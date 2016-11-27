using NPOIHelper.Contract;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Remoting;
using System.Runtime.Remoting.Channels;
using System.Runtime.Remoting.Channels.Tcp;
using System.Runtime.Serialization.Formatters;
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

            //RemotingConfiguration.Configure("NPOIHelper.Client.exe.config", false);
            RemotingConfiguration.Configure(AppDomain.CurrentDomain.SetupInformation.ConfigurationFile, false);
            proxy = (IPrint)Activator.GetObject(typeof(IPrint), "tcp://localhost:1235/Print/PrintURL");

            //message = Console.ReadLine();

            //ChannelServices.RegisterChannel(new TcpClientChannel(), true);
            //proxy = (IPrint)Activator.GetObject(typeof(IPrint), "tcp://localhost:1235/RemotePrint");
            Print();
        }

        private static void Print()
        {
            try
            {
                proxy.ExcelPrint("d:\\test.xls", "疑似黑广播信号出现情况1", new PrintCallBackHandler());
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
