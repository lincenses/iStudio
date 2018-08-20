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
            

            OpenFileDialog openDialog = new OpenFileDialog();
            openDialog.Filter = "*.xlsx|*.xlsx|*.xls|*.xls";
            if (openDialog.ShowDialog() == DialogResult.OK)
            {
                DateTime startTime = DateTime.Now;
                //DataSet dataSet = Studio.Extension.NPOIMethod.GetDataSetFromExcelFile(openDialog.FileName, true, true, false);
                //DataTable dataTable = Studio.Extension.NPOIMethod.GetDataTableFromExcelActiveSheet(openDialog.FileName, 1, 0);
                DataTable dataTable = Studio.Extension.NPOIMethod.GetDataTableFromExcelActiveSheet(openDialog.FileName, 1, 0, true, "№", "工号", "姓名", "生日", "年龄", "已婚", "备注", "","","","");
                string diff = DateTime.Now.Subtract(startTime).ToString();

                dataGridView1.DataSource = dataTable;

                MessageBox.Show(diff);
            }
            return;

            //SaveFileDialog dialog = new SaveFileDialog();
            //dialog.Filter = "*.xlsx|*.xlsx|*.xls|*.xls";
            //dialog.FileName = "ExcelColor.xlsx";
            //if (dialog.ShowDialog() == DialogResult.OK)
            //{
            //    DateTime startTime = DateTime.Now;
            //    Studio.Demo.NPOIDemo.CreateExcelColorFile(dialog.FileName);
            //    string time = DateTime.Now.Subtract(startTime).ToString();
            //    MessageBox.Show("创建成功！" + Environment.NewLine + time);
            //}
            //return;
        }
        #endregion
    }
}
