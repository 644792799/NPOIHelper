using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Runtime.Remoting;
using System.Runtime.Remoting.Channels;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace NPOIHelper.WFService
{
    public partial class WFService : Form
    {
        public WFService()
        {
            InitializeComponent();
        }

        private void WFService_Load(object sender, EventArgs e)
        {
            RemotingConfiguration.Configure(AppDomain.CurrentDomain.SetupInformation.ConfigurationFile, false);
            ;
            lstBxMessage.Items.Add("服务开启中......");
            var channels = ChannelServices.RegisteredChannels[0];
            lstBxMessage.Items.Add(RemotingConfiguration.GetRegisteredWellKnownServiceTypes()[0].ToString());
        }

        private void WFService_Resize(object sender, EventArgs e)
        {
            if (this.WindowState == FormWindowState.Minimized)    //最小化到系统托盘
            {
                notifyIcon1.Visible = true;    //显示托盘图标
                this.Hide();    //隐藏窗口
            }
        }

        private void WFService_FormClosing(object sender, FormClosingEventArgs e)
        {
            //注意判断关闭事件Reason来源于窗体按钮，否则用菜单退出时无法退出!
            if (e.CloseReason == CloseReason.UserClosing)
            {
                e.Cancel = true;    //取消"关闭窗口"事件
                this.WindowState = FormWindowState.Minimized;    //使关闭时窗口向右下角缩小的效果
                notifyIcon1.Visible = true;
                this.Hide();
                return;
            }
        }

        private void notifyIcon1_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            notifyIcon1.Visible = false;
            this.Show();
            WindowState = FormWindowState.Normal;
            this.Focus();
        }

        private void contextMenuStrip1_ItemClicked(object sender, ToolStripItemClickedEventArgs e)
        {
            switch (e.ClickedItem.Text)
            {
                case "显示窗体":
                    notifyIcon1.Visible = false;
                    this.Show();
                    WindowState = FormWindowState.Normal;
                    this.Focus();
                    break;
                case "退出":
                    DialogResult dr = MessageBox.Show("退出后将不能正常使用报表打印功能，您确定要退出NPOI打印服务吗?", "NPOI打印服务温馨提示", MessageBoxButtons.OKCancel, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2);
                    if (dr == System.Windows.Forms.DialogResult.OK)
                    {
                        Application.Exit();
                    }
                    break;
            }
        }
    }
}
