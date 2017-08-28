using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Studio.Monitor
{
    public static class DebugMonitor
    {
        #region 私有成员
        private static object _WriteSyncLock = new object();
        private static System.Diagnostics.Process _MonitorProcess;
        #endregion

        #region 公有属性

        #region 获取监视器标题
        public static string Title { get; private set; } = "";
        #endregion

        #region 获取监视器是否已运行
        public static bool IsRunning { get; private set; } = false;
        #endregion

        #region 获取监视器是否禁用关闭按钮
        public static bool DisableCloseButton { get; private set; } = false;
        #endregion

        #region 获取监视器背景色
        public static System.Drawing.Color BackColor { get; private set; } = System.Drawing.Color.FromArgb(50, 50, 50);
        #endregion

        #region 获取监视器前景色
        public static System.Drawing.Color ForeColor { get; private set; } = System.Drawing.Color.FromArgb(78, 210, 130);
        #endregion

        #region 获取监控器程序文件名
        public static string MonitorProgramFileName => System.Windows.Forms.Application.StartupPath + "\\StudioMonitor.exe";
        #endregion

        #endregion

        #region 构造函数
        static DebugMonitor()
        {
            Title = System.Windows.Forms.Application.ExecutablePath;
            _MonitorProcess = new System.Diagnostics.Process();
            _MonitorProcess.StartInfo.Arguments = GetStartArguments(Title);
            _MonitorProcess.StartInfo.FileName = MonitorProgramFileName;
            _MonitorProcess.StartInfo.UseShellExecute = false;
            _MonitorProcess.StartInfo.RedirectStandardInput = true;
            _MonitorProcess.EnableRaisingEvents = true;
            _MonitorProcess.Exited += new EventHandler((object sender, EventArgs e) => { IsRunning = false; });
        }
        #endregion

        #region 私有方法

        #region 获取参数字符串
        private static string GetStartArguments(string title)
        {
            List<string> args = new List<string>();
            args.Add("-host");
            args.Add(System.Diagnostics.Process.GetCurrentProcess().Id.ToString());
            args.Add("-n");
            args.Add("\"" + title + "\"");
            args.Add("-x");
            args.Add(DisableCloseButton.ToString());
            args.Add("-bc");
            args.Add(System.Drawing.ColorTranslator.ToHtml(BackColor));
            args.Add("-fc");
            args.Add(System.Drawing.ColorTranslator.ToHtml(ForeColor));
            return string.Join(" ", args);
        }
        #endregion

        #endregion

        #region 公有方法

        #region 显示监视器
        public static void Show()
        {
            if (IsRunning)
            {
                MonitorMethod.Activate(_MonitorProcess.MainWindowHandle);
            }
            else
            {
                //if (!System.IO.File.Exists(MonitorProgramFileName))
                //{ MonitorMethod.ReleaseMonitorFile(); }
                if (System.IO.File.Exists(MonitorProgramFileName))
                {
                    IsRunning = true;
                    _MonitorProcess.StartInfo.Arguments = GetStartArguments(Title);
                    _MonitorProcess.Start();
                }
            }
        }
        #endregion

        #region 关闭监视器
        public static void Close()
        {
            if (IsRunning)
            {
                try
                {
                    _MonitorProcess.Kill();
                    IsRunning = false;
                }
                catch { }
            }
        }
        #endregion

        #region 设置是否禁止关闭
        public static void SetDisableCloseButton(bool disable)
        {
            DisableCloseButton = disable;
            if (IsRunning)
            {
                lock (_WriteSyncLock)
                {
                    _MonitorProcess.StandardInput.WriteLine(string.Format("<disableClose>{0}</disableClose>", disable));
                }
            }
        }
        #endregion

        #region 设置背景色
        public static void SetBackColor(System.Drawing.Color backColor)
        {
            BackColor = backColor;
            if (IsRunning)
            {
                lock (_WriteSyncLock)
                {
                    _MonitorProcess.StandardInput.WriteLine(string.Format("<backColor>{0}</backColor>", System.Drawing.ColorTranslator.ToHtml(backColor)));
                }
            }
        }
        #endregion

        #region 设置前景色
        public static void SetForeColor(System.Drawing.Color foreColor)
        {
            ForeColor = foreColor;
            if (IsRunning)
            {
                lock (_WriteSyncLock)
                {
                    _MonitorProcess.StandardInput.WriteLine(string.Format("<foreColor>{0}</foreColor>", System.Drawing.ColorTranslator.ToHtml(foreColor)));
                }
            }
        }
        #endregion

        #region 向监视器写入消息
        public static void WriteLine(string format, params object[] args)
        {
            if (IsRunning)
            {
                lock (_WriteSyncLock)
                {
                    _MonitorProcess.StandardInput.WriteLine("<value>" + format + "</value>", args);
                }
            }
        }
        #endregion

        #region 向监视器写入消息
        public static void WriteLine(System.Drawing.Color fontColor, string format, params object[] args)
        {
            if (IsRunning)
            {
                lock (_WriteSyncLock)
                {
                    string color = string.Format("<color>{0}</color>", System.Drawing.ColorTranslator.ToHtml(fontColor));
                    _MonitorProcess.StandardInput.WriteLine(color + "<value>" + format + "</value>", args);
                }
            }
        }
        #endregion

        #endregion
    }
}
