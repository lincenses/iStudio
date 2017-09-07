using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;

namespace StudioDemo
{
    static class Program
    {
        /// <summary>
        /// 应用程序的主入口点。
        /// </summary>
        [STAThread]
        static void Main()
        {
            Studio.Configuration.DatabaseConfiguration configuration = new Studio.Configuration.DatabaseConfiguration();
            string cnnectionString = configuration.GetConnectionStringBuilder(30).ConnectionString;
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new FormMain());
        }
    }
}
