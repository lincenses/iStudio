using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Studio.Monitor
{
    internal static class MonitorMethod
    {
        #region 用于激活进程主窗体的API函数
        [System.Runtime.InteropServices.DllImport("user32.dll", CharSet = System.Runtime.InteropServices.CharSet.Auto, ExactSpelling = true)]
        private static extern bool SetForegroundWindow(IntPtr hWnd);

        [System.Runtime.InteropServices.DllImport("user32.dll", CharSet = System.Runtime.InteropServices.CharSet.Auto, ExactSpelling = true)]
        private static extern bool IsIconic(IntPtr hWnd);

        [System.Runtime.InteropServices.DllImport("user32.dll", CharSet = System.Runtime.InteropServices.CharSet.Auto, ExactSpelling = true)]
        public static extern bool ShowWindowAsync(IntPtr hWnd, int nCmdShow);
        #endregion

        #region 激活进程
        public static void Activate(IntPtr mainWindowHandle)
        {
            if (mainWindowHandle != IntPtr.Zero)
            {
                if (SetForegroundWindow(mainWindowHandle))
                {
                    if (IsIconic(mainWindowHandle))
                    {
                        int SW_RESTORE = 9;
                        ShowWindowAsync(mainWindowHandle, SW_RESTORE);
                    }
                    else
                    {
                        //int SW_SHOWNORMAL = 1;
                        //ShowWindowAsync(mainWindowHandle, SW_SHOWNORMAL);
                    }
                }
            }
        }
        #endregion

        #region 释放监视器程序
        public static void ReleaseMonitorFile()
        {
            System.Reflection.Assembly assembly = System.Reflection.Assembly.GetExecutingAssembly();
            string resourceName = assembly.GetManifestResourceNames().FirstOrDefault(x => x.EndsWith("StudioMonitor.exe"));
            if (!string.IsNullOrWhiteSpace(resourceName))
            {
                using (System.IO.Stream stream = assembly.GetManifestResourceStream(resourceName))
                {
                    if (stream != null)
                    {
                        byte[] buff = new byte[stream.Length];
                        stream.Read(buff, 0, buff.Length);
                        using (System.IO.FileStream fileStream = new System.IO.FileStream(System.Windows.Forms.Application.StartupPath + "\\StudioMonitor.exe", System.IO.FileMode.Create, System.IO.FileAccess.ReadWrite))
                        { fileStream.Write(buff, 0, buff.Length); }
                    }
                }
            }
        }
        #endregion
    }
}
