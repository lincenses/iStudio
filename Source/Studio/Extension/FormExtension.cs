using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Studio
{
    public static class FormExtension
    {
        #region 获取Mdi工作区
        public static System.Windows.Forms.MdiClient GetMdiClient(this System.Windows.Forms.Form mdiMainForm)
        {
            foreach (System.Windows.Forms.Control item in mdiMainForm.Controls)
            {
                System.Windows.Forms.MdiClient mdiClient = item as System.Windows.Forms.MdiClient;
                if (mdiClient != null)
                {
                    return mdiClient;
                }
            }
            return null;
        }
        #endregion

        #region 设置Mdi主窗体中Mdi工作区的边框
        public static void SetMdiContainerBorderStyle(this System.Windows.Forms.Form mdiMainForm, System.Windows.Forms.BorderStyle borderStyle)
        {
            System.Windows.Forms.MdiClient mdiClient = GetMdiClient(mdiMainForm);
            if (mdiClient != null)
            {
                int style = StudioNative.GetWindowLong(mdiClient.Handle, StudioNative.GWL_STYLE);
                int exStyle = StudioNative.GetWindowLong(mdiClient.Handle, StudioNative.GWL_EXSTYLE);
                if (borderStyle == System.Windows.Forms.BorderStyle.Fixed3D)
                {
                    style = 1442906112;
                    exStyle = 512;
                }
                else if (borderStyle == System.Windows.Forms.BorderStyle.FixedSingle)
                {
                    style |= StudioNative.WS_BORDER;
                    exStyle &= ~StudioNative.WS_EX_CLIENTEDGE;
                }
                else if (borderStyle == System.Windows.Forms.BorderStyle.None)
                {
                    style &= ~StudioNative.WS_BORDER;
                    exStyle &= ~StudioNative.WS_EX_CLIENTEDGE;
                }
                StudioNative.SetWindowLong(mdiClient.Handle, StudioNative.GWL_STYLE, style);
                StudioNative.SetWindowLong(mdiClient.Handle, StudioNative.GWL_EXSTYLE, exStyle);

                int flags = StudioNative.SWP_NOACTIVATE | StudioNative.SWP_NOMOVE | StudioNative.SWP_NOSIZE | StudioNative.SWP_NOZORDER | StudioNative.SWP_NOOWNERZORDER | StudioNative.SWP_FRAMECHANGED;
                StudioNative.SetWindowPos(mdiClient.Handle, IntPtr.Zero, 0, 0, 0, 0, flags);
            }
        }
        #endregion

        #region 独立显示Mdi子窗体
        public static void ShowMdiChildForm(this System.Windows.Forms.Form mainForm, System.Windows.Forms.Form childForm, bool displayControlBox = false)
        {
            foreach (System.Windows.Forms.Form item in mainForm.MdiChildren)
            {
                if (item.Name == childForm.Name)
                {
                    //item.Focus();
                    //item.WindowState = System.Windows.Forms.FormWindowState.Maximized;
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
                childForm.ControlBox = displayControlBox;
                childForm.MdiParent = mainForm;
                childForm.WindowState = System.Windows.Forms.FormWindowState.Maximized;
                childForm.Show();
            }

        }
        #endregion

        #region 禁用窗体关闭按钮
        public static void DisableCloseButton(this System.Windows.Forms.Form owner)
        {
            IntPtr systemMenuHandle = StudioNative.GetSystemMenu(owner.Handle, false);
            StudioNative.EnableMenuItem(systemMenuHandle, StudioNative.SC_CLOSE, StudioNative.MF_GRAYED);
        }
        #endregion

        #region 启用窗体关闭按钮
        public static void EnableCloseButton(this System.Windows.Forms.Form owner)
        {
            IntPtr systemMenuHandle = StudioNative.GetSystemMenu(owner.Handle, false);
            StudioNative.EnableMenuItem(systemMenuHandle, StudioNative.SC_CLOSE, StudioNative.MF_ENABLED);
        }
        #endregion

        #region 禁用窗体关闭按钮（重载CreateParams属性）
        // 将如下代码添加到窗体的类文件中。
        //protected override CreateParams CreateParams
        //{
        //    get
        //    {
        //        int CP_NOCLOSE_BUTTON = 0x200;
        //        CreateParams createParams = base.CreateParams;
        //        createParams.ClassStyle = createParams.ClassStyle | CP_NOCLOSE_BUTTON;
        //        return createParams;
        //    }
        //}
        #endregion

        #region 获取窗体中未包含在语言文件中的控件名称
        //public static string[] GetNotIncludedInTheLanguagePageControlNames(this System.Windows.Forms.Form owner, Studio.Language.LanguagePage languagePage)
        //{
        //    List<string> controlNameList = new List<string>();
        //    if (!languagePage.Contains(owner.Name))
        //    { controlNameList.Add(owner.Name); }
        //    System.Reflection.BindingFlags bindingFlags = System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance;
        //    System.Reflection.FieldInfo[] fieldInfos = owner.GetType().GetFields(bindingFlags);
        //    foreach (System.Reflection.FieldInfo fieldInfo in fieldInfos)
        //    {
        //        object control = fieldInfo.GetValue(owner);
        //        if (control != null)
        //        {
        //            System.Reflection.PropertyInfo controlNameInfo = control.GetType().GetProperty("Name");
        //            if (controlNameInfo != null)
        //            {
        //                string controlName = controlNameInfo.GetValue(control, null).ToString();
        //                if (!languagePage.Contains(controlName))
        //                { controlNameList.Add(controlName); }
        //            }
        //        }
        //    }
        //    return controlNameList.ToArray();
        //}
        #endregion
    }
}
