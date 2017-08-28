using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;

namespace Studio.Monitor
{
    public enum MonitorType
    {
        Normal = 0,
        File = 1
    }

    static class Program
    {
        // 监视器标题。
        public static string MonitorTitle = "Monitor";
        // 用于随主程序关闭。
        public static int HostID = 0;
        public static System.Threading.Thread HostMonitorProcess;
        // 监视器类型。
        public static MonitorType MonitorType = MonitorType.Normal;
        // 是否禁用关闭按钮。
        public static bool DisableCloseButton = false;
        // 监视器背景色。
        public static System.Drawing.Color BackColor = System.Drawing.Color.FromArgb(50, 50, 50);
        // 监视器前景色。
        public static System.Drawing.Color ForeColor = System.Drawing.Color.FromArgb(78, 210, 130);

        #region 构造函数
        static Program()
        {
            HostMonitorProcess = new System.Threading.Thread(new System.Threading.ParameterizedThreadStart(HostMonitorProcessMethod))
            {
                IsBackground = true,
                Priority = System.Threading.ThreadPriority.Normal,
            };
        }
        #endregion

        /// <summary>
        /// 应用程序的主入口点。
        /// </summary>
        [STAThread]
        static void Main()
        {
            List<string> args = Environment.GetCommandLineArgs().ToList();
            int index = 0;
            index = args.IndexOf("-host");
            if (index > 0 && args.Count > index + 1)
            {
                string value = args[index + 1];
                if (int.TryParse(value, out int hostID))
                {
                    HostMonitorProcess.Start(hostID);
                }
            }
            index = args.IndexOf("-n");
            if (index > 0 && args.Count > index + 1)
            {
                string value = args[index + 1];
                if (!string.IsNullOrWhiteSpace(value))
                {
                    MonitorTitle = value;
                    MonitorType = MonitorType.Normal;
                }
            }
            index = args.IndexOf("-f");
            if (index > 0 && args.Count > index + 1)
            {
                string value = args[index + 1];
                if (!string.IsNullOrWhiteSpace(value))
                {
                    MonitorTitle = value;
                    MonitorType = MonitorType.File;
                }
            }
            index = args.IndexOf("-x");
            if (index > 0 && args.Count > index + 1)
            {
                string value = args[index + 1];
                if (bool.TryParse(value, out bool disableCloseButton))
                {
                    DisableCloseButton = disableCloseButton;
                }
            }
            index = args.IndexOf("-bc");
            if (index > 0 && args.Count > index + 1)
            {
                string value = args[index + 1];
                if (!string.IsNullOrWhiteSpace(value))
                {
                    BackColor = System.Drawing.ColorTranslator.FromHtml(value);
                }
            }
            index = args.IndexOf("-fc");
            if (index > 0 && args.Count > index + 1)
            {
                string value = args[index + 1];
                if (!string.IsNullOrWhiteSpace(value))
                {
                    ForeColor = System.Drawing.ColorTranslator.FromHtml(value);
                }
            }

            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new FormMonitor());
        }

        #region 监视主程序的线程，用于随主程序关闭
        private static void HostMonitorProcessMethod(object hostID)
        {
            if (int.TryParse(hostID.ToString(), out int hostProcessID))
            {
                System.Diagnostics.Process hostProcess = System.Diagnostics.Process.GetProcessById(hostProcessID);
                if (hostProcess != null)
                {
                    while (true)
                    {
                        if (hostProcess.HasExited)
                        {
                            DisableCloseButton = false;
                            Application.Exit();
                            //Environment.Exit(0);
                        }
                    }
                }
            }
        }
        #endregion

    }
}
