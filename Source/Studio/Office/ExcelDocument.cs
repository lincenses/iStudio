using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Studio.Office
{
    public class ExcelDocument
    {
        [System.Runtime.InteropServices.DllImport("user32.dll", CharSet = System.Runtime.InteropServices.CharSet.Auto)]
        private static extern int GetWindowThreadProcessId(IntPtr hwnd, out int ID);

        #region 私有成员

        /// <summary>
        /// Excel对象。
        /// </summary>
        private Microsoft.Office.Interop.Excel.Application _Excel;

        /// <summary>
        /// 是否已被释放。
        /// </summary>
        private bool _IsDisposed = false;

        /// <summary>
        /// 进程ID。
        /// </summary>
        private int _ProcessID = 0;

        #endregion


    }
}
