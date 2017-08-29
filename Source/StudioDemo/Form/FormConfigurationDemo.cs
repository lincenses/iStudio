using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace StudioDemo
{
    public partial class FormConfigurationDemo : Form
    {
        public FormConfigurationDemo()
        {
            InitializeComponent();
            InitializeEvent();
        }

        #region 初始化事件
        private void InitializeEvent()
        {
            toolStripButton1.Click += ToolStripButton1_Click;
            toolStripButton2.Click += ToolStripButton2_Click;
        }

        
        #endregion

        #region 创建数据库配置文件
        private void ToolStripButton1_Click(object sender, EventArgs e)
        {
            SaveFileDialog dialog = new SaveFileDialog();
            dialog.InitialDirectory = Application.StartupPath + "\\Configuration";
            if (!System.IO.Directory.Exists(dialog.InitialDirectory))
            { System.IO.Directory.CreateDirectory(dialog.InitialDirectory); }
            dialog.FileName = "Database.cfg";
            dialog.Filter = "*.cfg|*.cfg";
            if (dialog.ShowDialog() == DialogResult.OK)
            {
                Studio.Configuration.DatabaseConfiguration configuration = new Studio.Configuration.DatabaseConfiguration();
                configuration.DataSource = ".\\SQL2012";
                configuration.InitialCatalog = "Sems";
                configuration.UserID = "sa";
                configuration.Password = "Admin123456";
                configuration.Save(dialog.FileName, true);
                webBrowser1.Url = new Uri(dialog.FileName);
            }
        }
        #endregion

        #region 读取数据库配置文件
        private void ToolStripButton2_Click(object sender, EventArgs e)
        {
            webBrowser1.Url = null;
        }
        #endregion
    }
}
