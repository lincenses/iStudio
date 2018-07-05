using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Studio.Extension
{
    public static class FormMethod
    {
        #region 独立显示Mdi子窗体
        public static void ShowMdiChildForm(this System.Windows.Forms.Form mainForm, System.Windows.Forms.Form childForm, bool showCloseButton)
        {
            if (mainForm.MainMenuStrip != null)
            {
                mainForm.MainMenuStrip.ItemAdded -= MainMenuStrip_ItemAdded;
                mainForm.MainMenuStrip.ItemAdded += MainMenuStrip_ItemAdded;
            }
            foreach (System.Windows.Forms.Form item in mainForm.MdiChildren)
            {
                if (item.Name == childForm.Name)
                {
                    item.Focus();
                    item.WindowState = System.Windows.Forms.FormWindowState.Maximized;
                    return;
                }
                else
                {
                    item.Close();
                    if (!item.IsDisposed)
                    { return; }
                }
            }
            if (!childForm.IsDisposed)
            {
                childForm.ControlBox = showCloseButton;
                childForm.MdiParent = mainForm;
                childForm.WindowState = System.Windows.Forms.FormWindowState.Maximized;
                childForm.Show();
            }
        }

        #region 主菜单的 ItemAdded 事件
        private static void MainMenuStrip_ItemAdded(object sender, System.Windows.Forms.ToolStripItemEventArgs e)
        {
            if (e.Item.Text.Length == 0 || e.Item.Text.Contains("(&R)") || e.Item.Text.Contains("(&N)"))
            { e.Item.Visible = false; }
            else
            { (sender as System.Windows.Forms.MenuStrip).Items[0].Image = e.Item.Image; }

        }
        #endregion

        #endregion

        #region 获取 Mdi 工作区
        private static System.Windows.Forms.MdiClient GetMdiClient(this System.Windows.Forms.Form mdiMainForm)
        {
            foreach (System.Windows.Forms.Control item in mdiMainForm.Controls)
            {
                if (item is System.Windows.Forms.MdiClient mdiClient)
                { return mdiClient; }
            }
            return null;
        }
        #endregion

        #region 设置 Mdi 工作区背景色
        public static void SetMdiContainerBackColor(this System.Windows.Forms.Form mdiMainForm, System.Drawing.Color backColor)
        {
            System.Windows.Forms.MdiClient mdiClient = GetMdiClient(mdiMainForm);
            if (mdiClient != null)
            { mdiClient.BackColor = backColor; }
        }
        #endregion

        #region 设置 Mdi 工作区边框
        public static void SetMdiContainerBorderStyle(this System.Windows.Forms.Form mdiMainForm, System.Windows.Forms.BorderStyle borderStyle)
        {
            System.Windows.Forms.MdiClient mdiClient = GetMdiClient(mdiMainForm);
            if (mdiClient != null)
            {
                int style = GetWindowLong(mdiClient.Handle, GWL_STYLE);
                int exStyle = GetWindowLong(mdiClient.Handle, GWL_EXSTYLE);
                if (borderStyle == System.Windows.Forms.BorderStyle.Fixed3D)
                {
                    style = 1442906112;
                    exStyle = 512;
                }
                else if (borderStyle == System.Windows.Forms.BorderStyle.FixedSingle)
                {
                    style |= WS_BORDER;
                    exStyle &= ~WS_EX_CLIENTEDGE;
                }
                else if (borderStyle == System.Windows.Forms.BorderStyle.None)
                {
                    style &= ~WS_BORDER;
                    exStyle &= ~WS_EX_CLIENTEDGE;
                }
                SetWindowLong(mdiClient.Handle, GWL_STYLE, style);
                SetWindowLong(mdiClient.Handle, GWL_EXSTYLE, exStyle);

                int flags = SWP_NOACTIVATE | SWP_NOMOVE | SWP_NOSIZE | SWP_NOZORDER | SWP_NOOWNERZORDER | SWP_FRAMECHANGED;
                SetWindowPos(mdiClient.Handle, IntPtr.Zero, 0, 0, 0, 0, flags);
            }
        }

        #region API 函数
        private const int GWL_EXSTYLE = -20;
        private const int GWL_STYLE = -16;

        private const int WS_BORDER = 0x00800000;
        private const int WS_EX_CLIENTEDGE = 0x00000200;

        [System.Runtime.InteropServices.DllImport("user32.dll", EntryPoint = "GetWindowLong", CharSet = System.Runtime.InteropServices.CharSet.Auto)]
        private static extern int GetWindowLong32(IntPtr hWnd, int nIndex);

        [System.Runtime.InteropServices.DllImport("user32.dll", EntryPoint = "GetWindowLongPtr", CharSet = System.Runtime.InteropServices.CharSet.Auto)]
        private static extern int GetWindowLongPtr64(IntPtr hWnd, int nIndex);

        private static int GetWindowLong(IntPtr hWnd, int nIndex)
        {
            if (IntPtr.Size == 4)
            {
                return GetWindowLong32(hWnd, nIndex);
            }
            return GetWindowLongPtr64(hWnd, nIndex);
        }

        [System.Runtime.InteropServices.DllImport("user32.dll", EntryPoint = "SetWindowLong", CharSet = System.Runtime.InteropServices.CharSet.Auto)]
        private static extern int SetWindowLongPtr32(IntPtr hWnd, int nIndex, int dwNewLong);
        [System.Runtime.InteropServices.DllImport("user32.dll", EntryPoint = "SetWindowLongPtr", CharSet = System.Runtime.InteropServices.CharSet.Auto)]
        private static extern int SetWindowLongPtr64(IntPtr hWnd, int nIndex, int dwNewLong);

        private static int SetWindowLong(IntPtr hWnd, int nIndex, int dwNewLong)
        {
            if (IntPtr.Size == 4)
            {
                return SetWindowLongPtr32(hWnd, nIndex, dwNewLong);
            }
            return SetWindowLongPtr64(hWnd, nIndex, dwNewLong);
        }

        private const int SWP_NOSIZE = 0x0001;//
        private const int SWP_NOMOVE = 0x0002;//
        private const int SWP_NOZORDER = 0x0004;//
        private const int SWP_NOREDRAW = 0x0008;
        private const int SWP_NOACTIVATE = 0x0010;//
        private const int SWP_FRAMECHANGED = 0x0020;//
        private const int SWP_SHOWWINDOW = 0x0040;
        private const int SWP_HIDEWINDOW = 0x0080;
        private const int SWP_NOCOPYBITS = 0x0100;
        private const int SWP_NOOWNERZORDER = 0x0200;//
        private const int SWP_NOSENDCHANGING = 0x0400;

        [System.Runtime.InteropServices.DllImport("user32.dll", CharSet = System.Runtime.InteropServices.CharSet.Auto, ExactSpelling = true)]
        private static extern bool SetWindowPos(IntPtr hWnd, IntPtr hWndInsertAfter, int x, int y, int cx, int cy, int flags);

        #endregion



        #endregion

    }
}
