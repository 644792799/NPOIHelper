using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Remoting;
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
        }
    }
}
