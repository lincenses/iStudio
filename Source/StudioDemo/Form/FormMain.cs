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
            this.GetMdiClient().BackColor = Color.FromKnownColor(KnownColor.Window);
            this.SetMdiContainerBorderStyle(BorderStyle.None);
            menuStrip1.DisableMinMaxBox();
        }
        #endregion

        #region 初始化事件
        private void InitializeEvent()
        {
            toolStripMenuItemConfiguration.Click += ToolStripMenuItemConfiguration_Click;
        }
        #endregion

        #region 配置文件演示
        private void ToolStripMenuItemConfiguration_Click(object sender, EventArgs e)
        {
            this.ShowMdiChildForm(new FormConfigurationDemo(), true);
        }
        #endregion
    }
}
