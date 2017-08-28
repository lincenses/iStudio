using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Studio
{
    public static class StudioNative
    {
        [System.Runtime.InteropServices.DllImport("kernel32.dll")]
        public static extern Boolean AllocConsole();



        [System.Runtime.InteropServices.DllImport("kernel32.dll")]
        public static extern bool FreeConsole();



        public const int WS_BORDER = 0x00800000;
        public const int WS_EX_CLIENTEDGE = 0x00000200;



        public const int GWL_EXSTYLE = -20;
        public const int GWL_STYLE = -16;
        [System.Runtime.InteropServices.DllImport("user32.dll", EntryPoint = "GetWindowLong", CharSet = System.Runtime.InteropServices.CharSet.Auto)]
        public static extern int GetWindowLong32(IntPtr hWnd, int nIndex);

        [System.Runtime.InteropServices.DllImport("user32.dll", EntryPoint = "GetWindowLongPtr", CharSet = System.Runtime.InteropServices.CharSet.Auto)]
        public static extern int GetWindowLongPtr64(IntPtr hWnd, int nIndex);

        public static int GetWindowLong(IntPtr hWnd, int nIndex)
        {
            if (IntPtr.Size == 4)
            {
                return GetWindowLong32(hWnd, nIndex);
            }
            return GetWindowLongPtr64(hWnd, nIndex);
        }



        [System.Runtime.InteropServices.DllImport("user32.dll", EntryPoint = "SetWindowLong", CharSet = System.Runtime.InteropServices.CharSet.Auto)]
        public static extern int SetWindowLongPtr32(IntPtr hWnd, int nIndex, int dwNewLong);
        [System.Runtime.InteropServices.DllImport("user32.dll", EntryPoint = "SetWindowLongPtr", CharSet = System.Runtime.InteropServices.CharSet.Auto)]
        public static extern int SetWindowLongPtr64(IntPtr hWnd, int nIndex, int dwNewLong);

        public static int SetWindowLong(IntPtr hWnd, int nIndex, int dwNewLong)
        {
            if (IntPtr.Size == 4)
            {
                return SetWindowLongPtr32(hWnd, nIndex, dwNewLong);
            }
            return SetWindowLongPtr64(hWnd, nIndex, dwNewLong);
        }



        public const int SWP_NOSIZE = 0x0001;//
        public const int SWP_NOMOVE = 0x0002;//
        public const int SWP_NOZORDER = 0x0004;//
        public const int SWP_NOREDRAW = 0x0008;
        public const int SWP_NOACTIVATE = 0x0010;//
        public const int SWP_FRAMECHANGED = 0x0020;//
        public const int SWP_SHOWWINDOW = 0x0040;
        public const int SWP_HIDEWINDOW = 0x0080;
        public const int SWP_NOCOPYBITS = 0x0100;
        public const int SWP_NOOWNERZORDER = 0x0200;//
        public const int SWP_NOSENDCHANGING = 0x0400;
        [System.Runtime.InteropServices.DllImport("user32.dll", CharSet = System.Runtime.InteropServices.CharSet.Auto, ExactSpelling = true)]
        public static extern bool SetWindowPos(IntPtr hWnd, IntPtr hWndInsertAfter, int x, int y, int cx, int cy, int flags);




        [System.Runtime.InteropServices.DllImport("user32.dll", CharSet = System.Runtime.InteropServices.CharSet.Auto, ExactSpelling = true)]
        public static extern IntPtr GetSystemMenu(IntPtr hWnd, bool bRevert);



        public const int SC_CLOSE = 0xF060;
        public const int MF_ENABLED = 0x00000000;
        public const int MF_GRAYED = 0x00000001;
        public const int MF_DISABLED = 0x00000002;
        [System.Runtime.InteropServices.DllImport("user32.dll", CharSet = System.Runtime.InteropServices.CharSet.Auto, ExactSpelling = true)]
        public static extern bool EnableMenuItem(IntPtr hMenu, int UIDEnabledItem, int uEnable);

    }
}
