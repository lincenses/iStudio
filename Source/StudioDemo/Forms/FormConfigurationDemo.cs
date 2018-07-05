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
            InitializeForm();
            InitializeEvent();
        }

        #region 初始化窗体
        private void InitializeForm()
        {
            //webBrowser1.Visible = false;
        }
        #endregion

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
            for (int i = 0; i < panel1.Controls.Count; i++)
            {
                panel1.Controls.RemoveAt(i);
            }

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
                WebBrowser webBorwser = new WebBrowser();
                webBorwser.Name = "webBorwserTemp";
                webBorwser.Dock = DockStyle.Fill;
                webBorwser.Url = new Uri(dialog.FileName);
                panel1.Controls.Add(webBorwser);
            }
        }
        #endregion

        #region 读取数据库配置文件
        private void ToolStripButton2_Click(object sender, EventArgs e)
        {
            for (int i = 0; i < panel1.Controls.Count; i++)
            {
                panel1.Controls.RemoveAt(i);
            }
            OpenFileDialog dialog = new OpenFileDialog();
            dialog.FileName = "Database.cfg";
            dialog.Filter = "*.cfg|*.cfg";
            if (dialog.ShowDialog() == DialogResult.OK)
            {
                Studio.Configuration.DatabaseConfiguration configuration = Studio.Configuration.DatabaseConfiguration.LoadFrom(dialog.FileName);
                RichTextBox richTextBox = new RichTextBox();
                richTextBox.Name = "richTextBoxTemp";
                richTextBox.Dock = DockStyle.Fill;
                richTextBox.Font = new Font("Consolas", 14, FontStyle.Regular);
                richTextBox.AppendText(string.Format("DataSource:\t\t{0}" + Environment.NewLine, configuration.DataSource));
                richTextBox.AppendText(string.Format("InitialCatalog:\t{0}" + Environment.NewLine, configuration.InitialCatalog));
                richTextBox.AppendText(string.Format("UserID:\t\t\t{0}" + Environment.NewLine, configuration.UserID));
                richTextBox.AppendText(string.Format("Password:\t\t\t{0}" + Environment.NewLine, configuration.Password));
                
                panel1.Controls.Add(richTextBox);
            }
        }
        #endregion
    }
}
