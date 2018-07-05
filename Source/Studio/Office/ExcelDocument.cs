using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Studio.Office
{
    public class ExcelDocument : IDisposable
    {
        [System.Runtime.InteropServices.DllImport("user32.dll", CharSet = System.Runtime.InteropServices.CharSet.Auto)]
        private static extern int GetWindowThreadProcessId(IntPtr hwnd, out int ID);

        #region 私有成员

        private Microsoft.Office.Interop.Excel.Application _Excel;

        private bool _IsDisposed = false;

        private int _ProcessID = 0;

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
                _Excel.Application.DisplayAlerts = false;
                _Excel.Workbooks.Close();
                _Excel.Application.DisplayAlerts = true;
                _Excel.Quit();
                if (_ProcessID != 0)
                { System.Diagnostics.Process.GetProcessById(_ProcessID).Kill(); }
            }
            _IsDisposed = true; // 标识此对象已释放
        }
        #endregion

    }
}
