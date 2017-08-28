using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Studio
{
    public static class MenuStripExtension
    {
        #region 禁止菜单显示最小化和最大化按钮
        public static void DisableMinMaxBox(this System.Windows.Forms.MenuStrip owner)
        {
            owner.ItemAdded += new System.Windows.Forms.ToolStripItemEventHandler((object sender, System.Windows.Forms.ToolStripItemEventArgs e) =>
            {
                if (e.Item.Text.Length == 0 || e.Item.Text.Contains("(&R)") || e.Item.Text.Contains("(&N)"))
                {
                    e.Item.Visible = false;
                }
                else
                {
                    (sender as System.Windows.Forms.MenuStrip).Items[0].Image = e.Item.Image;
                }
            });
        }
        #endregion

        #region 初始化语言菜单
        //public static void InitializeLanguageMenuItem(this System.Windows.Forms.ToolStripMenuItem languageMenuItem, System.Windows.Forms.Form mainForm)
        //{
        //    if (languageMenuItem.HasDropDownItems)
        //    { languageMenuItem.DropDownItems.Clear(); }
        //    foreach (System.Data.DataRow row in Studio.Language.LanguageManager.Dictionaries.CultureNameTable.Rows)
        //    {
        //        string nativeName = row["NativeName"].ToString();
        //        string cultureName = row["CultureName"].ToString();
        //        System.Windows.Forms.ToolStripMenuItem menuItem = new System.Windows.Forms.ToolStripMenuItem();
        //        menuItem.Name = "toolStripMenuItem" + cultureName.Replace("-", "");
        //        menuItem.Text = nativeName;
        //        if (cultureName == Studio.Language.LanguageManager.CultureName)
        //        { menuItem.Checked = true; }
        //        languageMenuItem.DropDownItems.Add(menuItem);
        //        menuItem.Click += new EventHandler((object sender, EventArgs e) =>
        //        {
        //            foreach (System.Windows.Forms.ToolStripMenuItem item in languageMenuItem.DropDownItems)
        //            { item.Checked = false; }
        //            menuItem.Checked = true;
        //            Studio.Language.LanguageManager.SetCultureName(cultureName);
        //            Studio.Language.ILanguage iLanguage = mainForm as Studio.Language.ILanguage;
        //            if (iLanguage != null)
        //            { iLanguage.InitializeLanguage(); }
        //        });
        //    }
        //}
        #endregion
    }
}
