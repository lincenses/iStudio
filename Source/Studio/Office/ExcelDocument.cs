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


    }
}
