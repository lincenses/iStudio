using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Studio.Office.Excel
{
    public class ExcelDocument : IDisposable
    {
        #region API函数
        [System.Runtime.InteropServices.DllImport("user32.dll", CharSet = System.Runtime.InteropServices.CharSet.Auto)]
        private static extern int GetWindowThreadProcessId(IntPtr hwnd, out int ID);
        #endregion

        #region 私有成员

        private Microsoft.Office.Interop.Excel.Application _XlApplication = null;

        private bool _IsDisposed = false;

        private int _ProcessID = 0;

        private ExcelSheets _ExcelSheets = null;

        public ExcelSheets Sheets { get => _ExcelSheets; }

        #endregion

        #region 实现接口
        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        protected virtual void Dispose(bool disposing)
        {
            if (!_IsDisposed)// 如果资源未释放 这个判断主要用了防止对象被多次释放
            {
                if (disposing)
                {
                    // 释放托管资源
                }
                // 释放非托管资源
                _XlApplication.Application.DisplayAlerts = false;
                _XlApplication.Workbooks.Close();
                _XlApplication.Application.DisplayAlerts = true;
                _XlApplication.Quit();
                if (_ProcessID != 0)
                { System.Diagnostics.Process.GetProcessById(_ProcessID).Kill(); }
            }
            _IsDisposed = true; // 标识此对象已释放
        }
        #endregion

        #region 构造函数

        #region 初始化此类的新实例
        public ExcelDocument()
        {
            _XlApplication = new Microsoft.Office.Interop.Excel.Application();
            GetWindowThreadProcessId(new IntPtr(_XlApplication.Hwnd), out _ProcessID);
            _XlApplication.Application.DisplayAlerts = false;
            _XlApplication.Workbooks.Add(Type.Missing);

            //_ExcelSheets = new ExcelSheets(_XlApplication.Worksheets);
            //ExcelSheet excelSheet = new ExcelSheet(_XlApplication.ActiveSheet, _ExcelSheets);
            //_ExcelSheets.ExcelSheetList.Add(excelSheet);

        }

        public ExcelDocument(string templateFileName, bool includeHide = false)
        {
            _XlApplication = new Microsoft.Office.Interop.Excel.Application();
            GetWindowThreadProcessId(new IntPtr(_XlApplication.Hwnd), out _ProcessID);
            _XlApplication.Application.DisplayAlerts = false;
            _XlApplication.Workbooks.Add(new System.IO.FileInfo(templateFileName).FullName);

            //_ExcelSheets = new ExcelSheets(_XlApplication.Worksheets);
            //for (int i = 1; i < _XlApplication.Worksheets.Count + 1; i++)
            //{
            //    Microsoft.Office.Interop.Excel.Worksheet xlSheet = _XlApplication.Worksheets[i];
            //    if (includeHide == false && xlSheet.Visible != Microsoft.Office.Interop.Excel.XlSheetVisibility.xlSheetVisible)
            //    { continue; }
            //    ExcelSheet excelSheet = new ExcelSheet(xlSheet, _ExcelSheets);
            //    _ExcelSheets.ExcelSheetList.Add(excelSheet);
            //}
        }
        #endregion

        #region 析构函数
        ~ExcelDocument()
        {
            Dispose(false);
        }
        #endregion

        #endregion

        #region 共有方法

        #region 关闭Excel
        public void Close()
        {
            Dispose();
        }
        #endregion

        #region 保存Excel文件（XLS格式）
        public void SaveAsXLS(string fileName)
        {
            string fullName = new System.IO.FileInfo(fileName).FullName;
            _XlApplication.ActiveWorkbook.SaveAs(fullName, Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookNormal, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);

        }
        #endregion

        #region 保存Excel文件（XLS格式）
        public void SaveAsXLSX(string fileName)
        {
            string fullName = new System.IO.FileInfo(fileName).FullName;
            _XlApplication.ActiveWorkbook.SaveAs(fullName, Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookDefault, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);

        }
        #endregion

        #region 保存Excel文件（PDF格式）
        public void SaveAsPDF(string fileName)
        {
            string fullName = new System.IO.FileInfo(fileName).FullName;
            _XlApplication.ActiveWorkbook.ExportAsFixedFormat(Microsoft.Office.Interop.Excel.XlFixedFormatType.xlTypePDF, fullName, Microsoft.Office.Interop.Excel.XlFixedFormatQuality.xlQualityStandard, true, false, Type.Missing, Type.Missing, false, Type.Missing);

        }
        #endregion

        #endregion


    }
}
