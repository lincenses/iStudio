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

        #region 创建数据库配置文件
        private void ToolStripButton1_Click(object sender, EventArgs e)
        {
            Studio.Office.ExcelDocument excelDoc = new Studio.Office.ExcelDocument("zzz.xlsx");
            excelDoc.SetActiveSheet(1);
            excelDoc.SetCellValue(2, 2, 123456);
            excelDoc.SetActiveCell(2, 2);

            object value = excelDoc.GetCellValue();


            excelDoc.SaveAs("zzzzzz");
            excelDoc.Close();
        }
        #endregion
    }
}
