using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Linq;
using System.Runtime.Remoting;
using System.ServiceProcess;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace NPOIHelper.WinService
{
    public partial class Service1 : ServiceBase
    {
        volatile bool isstart = true; 

        public Service1()
        {
            InitializeComponent();
        }

        protected override void OnStart(string[] args)
        {
            RemotingConfiguration.Configure(AppDomain.CurrentDomain.SetupInformation.ConfigurationFile, false);
            //Thread t = new Thread(delegate()
            //{
            //    ThreadMethod();
            //});
            //t.Start();
        }

        protected override void OnStop()
        {
            isstart = false;
        }

        void ThreadMethod()
        {
            while (isstart)
            {
                try
                {
                    Thread.Sleep(30000);
                }
                catch { }
            }
        }
    }
}
