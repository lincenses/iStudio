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
    public partial class FormExcelDemo : Form
    {
        public FormExcelDemo()
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
            
        }
        #endregion

        #region Excel Demo
        private void ToolStripButton1_Click(object sender, EventArgs e)
        {
            SaveFileDialog dialog = new SaveFileDialog();
            dialog.Filter = "*.xlsx|*.xlsx|*.xls|*.xls";
            dialog.FileName = "ExcelColor.xlsx";
            if (dialog.ShowDialog() == DialogResult.OK)
            {
                Studio.Demo.NPOIDemo.NPOIMethod.CreateExcelColorFile(dialog.FileName);
                MessageBox.Show("创建成功！");
            }
            return;
        }
        #endregion
    }
}
