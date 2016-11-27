using NPOIHelper.Remoting;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Remoting;
using System.Runtime.Remoting.Channels;
using System.Runtime.Remoting.Channels.Tcp;
using System.Text;
using System.Threading.Tasks;

namespace NPOIHelper.ConsoleService
{
    public class Program
    {
        static void Main(string[] args)
        {
            RemotingConfiguration.Configure(AppDomain.CurrentDomain.SetupInformation.ConfigurationFile, false);
            Console.WriteLine("Remoting service starting....");
            Console.Read();

            //TcpServerChannel channel = new TcpServerChannel(1235);
            //ChannelServices.RegisterChannel(channel, true);
            //RemotingConfiguration.RegisterWellKnownServiceType(typeof(RemotePrint),
            //    "RemotePrint", WellKnownObjectMode.Singleton);
            //System.Console.WriteLine("Remoting service starting....");
            //System.Console.ReadLine();
        }
    }
}
