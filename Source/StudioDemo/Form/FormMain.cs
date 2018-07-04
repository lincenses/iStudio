using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Studio;

namespace StudioDemo
{
    public partial class FormMain : Form
    {
        public FormMain()
        {
            InitializeComponent();
            InitializeForm();
            InitializeEvent();
        }

        protected override CreateParams CreateParams
        {
            get
            {
                CreateParams createParams = base.CreateParams;
                createParams.ExStyle |= 0x02000000;
                //createParams.Style ^= 0x00C00000; //WS_CAPTION
                return createParams;
            }
        }

        #region 初始化窗体
        private void InitializeForm()
        {
            this.GetMdiClient().BackColor = Color.FromKnownColor(KnownColor.Window);
            this.SetMdiContainerBorderStyle(BorderStyle.None);
            menuStripMain.DisableMinMaxBox();
        }
        #endregion

        #region 初始化事件
        private void InitializeEvent()
        {
            toolStripMenuItemConfiguration.Click += ToolStripMenuItemConfiguration_Click;
            toolStripMenuItemMonitor.Click += ToolStripMenuItemMonitor_Click;
            
        }
        #endregion

        #region 配置文件演示
        private void ToolStripMenuItemConfiguration_Click(object sender, EventArgs e)
        {
            this.ShowMdiChildForm(new FormConfigurationDemo(), true);
        }
        #endregion

        #region 监视器演示
        private void ToolStripMenuItemMonitor_Click(object sender, EventArgs e)
        {
            this.ShowMdiChildForm(new FormDebugMonitorDemo(), true);
        }
        #endregion
    }
}
