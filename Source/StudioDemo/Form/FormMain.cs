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

        #region 初始化窗体
        private void InitializeForm()
        {
            Studio.Extension.FormMethod.SetMdiContainerBackColor(this, Color.FromKnownColor(KnownColor.ControlLight));
            Studio.Extension.FormMethod.SetMdiContainerBorderStyle(this, BorderStyle.FixedSingle);
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
            Studio.Extension.FormMethod.ShowMdiChildForm(this, new FormConfigurationDemo(), true);

        }
        #endregion

        #region 监视器演示
        private void ToolStripMenuItemMonitor_Click(object sender, EventArgs e)
        {
            Studio.Extension.FormMethod.ShowMdiChildForm(this, new FormDebugMonitorDemo(), true);

        }
        #endregion
    }
}
