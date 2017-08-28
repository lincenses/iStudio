using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace Studio.Monitor
{
    public partial class FormMonitor : Form
    {
        #region 私有成员
        private int _LastCount = 0;
        private DateTime _LastTime = DateTime.MinValue;

        private MonitorRichTextBox _MonitorRichTextBox;
        private System.Threading.Thread _NormalMonitorProcess;
        private System.Threading.Thread _FileMonitorProcess;
        #endregion

        #region 构造函数
        public FormMonitor()
        {
            InitializeComponent();
            InitializeForm();
            InitializeEvent();

        }
        #endregion

        #region 初始化窗体
        private void InitializeForm()
        {
            Text = Program.MonitorTitle;
            Size = new Size(640, 400);
            System.Reflection.Assembly assembly = System.Reflection.Assembly.GetExecutingAssembly();
            string iconResourceName = assembly.GetManifestResourceNames().FirstOrDefault(x => x.EndsWith("Program.ico"));
            if (!string.IsNullOrWhiteSpace(iconResourceName))
            {
                Icon = new Icon(assembly.GetManifestResourceStream(iconResourceName));
            }
            _MonitorRichTextBox = new MonitorRichTextBox()
            {
                ReadOnly = true,
                Dock = DockStyle.Fill,
                BorderStyle = BorderStyle.None,
                Font = new Font("Consolas", 11),
                BackColor = Program.BackColor,
                ForeColor = Program.ForeColor,
                ScrollBars = RichTextBoxScrollBars.Both,
                LanguageOption = RichTextBoxLanguageOptions.DualFont,
            };
            Controls.Add(_MonitorRichTextBox);

            StartNormalMonitorProcess();

            if (Program.MonitorType == MonitorType.File)
            {
                StartFileMonitorProcess();
            }
        }
        #endregion

        #region 初始化事件
        private void InitializeEvent()
        {
            FormClosing += new FormClosingEventHandler((object sender, FormClosingEventArgs e) =>
            {
                e.Cancel = Program.DisableCloseButton;
                if (Program.DisableCloseButton)
                {
                    return;
                }
                if (_NormalMonitorProcess != null)
                {
                    _NormalMonitorProcess.Abort();
                }
                if (_FileMonitorProcess != null)
                {
                    _FileMonitorProcess.Abort();
                }
            });
        }
        #endregion

        #region 初始化并启动常规监控线程
        private void StartNormalMonitorProcess()
        {
            _NormalMonitorProcess = new System.Threading.Thread(new System.Threading.ThreadStart(NormalMonitorMethod))
            {
                IsBackground = true,
                Priority = System.Threading.ThreadPriority.Normal
            };
            _NormalMonitorProcess.Start();
        }

        private void NormalMonitorMethod()
        {
            System.Threading.Thread.Sleep(60);
            while (true)
            {
                string message = Console.ReadLine();
                if (!string.IsNullOrWhiteSpace(message))
                {
                    string value = System.Text.RegularExpressions.Regex.Match(message, @"(?<=<value>).*?(?=</value>)", System.Text.RegularExpressions.RegexOptions.IgnoreCase).Value;
                    string color = System.Text.RegularExpressions.Regex.Match(message, @"(?<=<color>).*?(?=</color>)", System.Text.RegularExpressions.RegexOptions.IgnoreCase).Value;
                    string foreColor = System.Text.RegularExpressions.Regex.Match(message, @"(?<=<foreColor>).*?(?=</foreColor>)", System.Text.RegularExpressions.RegexOptions.IgnoreCase).Value;
                    string backColor = System.Text.RegularExpressions.Regex.Match(message, @"(?<=<backColor>).*?(?=</backColor>)", System.Text.RegularExpressions.RegexOptions.IgnoreCase).Value;
                    string disableClose = System.Text.RegularExpressions.Regex.Match(message, @"(?<=<disableClose>).*?(?=</disableClose>)", System.Text.RegularExpressions.RegexOptions.IgnoreCase).Value;
                    string fileName = System.Text.RegularExpressions.Regex.Match(message, @"(?<=<fileName>).*?(?=</fileName>)", System.Text.RegularExpressions.RegexOptions.IgnoreCase).Value;
                    if (!string.IsNullOrWhiteSpace(disableClose))
                    {
                        if (bool.TryParse(disableClose, out bool result))
                        {
                            Program.DisableCloseButton = result;
                        }
                    }
                    if (!string.IsNullOrWhiteSpace(foreColor))
                    {
                        SetForeColor(System.Drawing.ColorTranslator.FromHtml(foreColor));
                    }
                    if (!string.IsNullOrWhiteSpace(backColor))
                    {
                        SetBackColor(System.Drawing.ColorTranslator.FromHtml(backColor));
                    }
                    if (!string.IsNullOrWhiteSpace(fileName))
                    {
                        SetFileName(fileName);
                    }
                    if (!string.IsNullOrWhiteSpace(value))
                    {
                        if (string.IsNullOrWhiteSpace(color))
                        {
                            DisplayMessage(value);
                        }
                        else
                        {
                            DisplayMessage(value, System.Drawing.ColorTranslator.FromHtml(color));
                        }
                    }
                }
            }
        }
        #endregion

        #region 初始化并启动文件监控线程
        private void StartFileMonitorProcess()
        {
            _FileMonitorProcess = new System.Threading.Thread(new System.Threading.ThreadStart(FileMonitorMethod))
            {
                IsBackground = true,
                Priority = System.Threading.ThreadPriority.Normal
            };
            _FileMonitorProcess.Start();
        }

        private void FileMonitorMethod()
        {
            System.Threading.Thread.Sleep(60);
            while (true)
            {
                if (System.IO.File.Exists(Program.MonitorTitle))
                {
                    DateTime fileTime = System.IO.File.GetLastWriteTime(Program.MonitorTitle);
                    if (fileTime.CompareTo(_LastTime) > 0)
                    {
                        int fileCount = 0;
                        List<string> fileValues = new List<string>();
                        try
                        {
                            fileValues = System.IO.File.ReadAllLines(Program.MonitorTitle).ToList();
                        }
                        catch { continue; }
                        fileCount = fileValues.Count;
                        if (fileCount > 0 && fileCount > _LastCount)
                        {
                            fileValues.RemoveRange(0, _LastCount);
                            DisplayMessage(string.Join(Environment.NewLine, fileValues));
                        }
                        _LastTime = fileTime;
                        _LastCount = fileCount;
                    }
                }
                else
                {
                    DisplayMessage(Program.MonitorTitle);
                    DisplayMessage("指定的文件不存在。");
                    System.Threading.Thread.Sleep(1000);
                }
            }
        }
        #endregion

        #region 设置前景色
        private void SetForeColor(System.Drawing.Color foreColor)
        {
            Program.ForeColor = foreColor;
            if (_MonitorRichTextBox.InvokeRequired)
            {
                _MonitorRichTextBox.Invoke(new MethodInvoker(() =>
                {
                    _MonitorRichTextBox.ForeColor = foreColor;
                }));
            }
            else
            {
                _MonitorRichTextBox.ForeColor = foreColor;
            }
        }
        #endregion

        #region 设置背景色
        private void SetBackColor(System.Drawing.Color backColor)
        {
            Program.BackColor = backColor;
            if (_MonitorRichTextBox.InvokeRequired)
            {
                _MonitorRichTextBox.Invoke(new MethodInvoker(() =>
                {
                    _MonitorRichTextBox.BackColor = backColor;
                }));
            }
            else
            {
                _MonitorRichTextBox.BackColor = backColor;
            }
        }
        #endregion

        #region 设置文件名
        private void SetFileName(string fileName)
        {
            if (InvokeRequired)
            {
                Invoke(new MethodInvoker(() =>
                {
                    Program.MonitorTitle = fileName;
                    _LastTime = DateTime.MinValue;
                    _LastCount = 0;
                    Text = Program.MonitorTitle;
                    _MonitorRichTextBox.Clear();
                }));
            }
            else
            {
                Program.MonitorTitle = fileName;
                _LastTime = DateTime.MinValue;
                _LastCount = 0;
                Text = Program.MonitorTitle;
                _MonitorRichTextBox.Clear();
            }
        }
        #endregion

        #region 显示消息
        private void DisplayMessage(string message)
        {
            if (_MonitorRichTextBox.InvokeRequired)
            {
                _MonitorRichTextBox.Invoke(new MethodInvoker(() =>
                {
                    if (message.Length + Environment.NewLine.Length + _MonitorRichTextBox.TextLength >= 4194304)
                    {
                        _MonitorRichTextBox.Clear();
                    }
                    _MonitorRichTextBox.AppendText(message + Environment.NewLine);
                    _MonitorRichTextBox.ScrollToCaret();
                }));
            }
            else
            {
                if (message.Length + Environment.NewLine.Length + _MonitorRichTextBox.TextLength >= 4194304)
                {
                    _MonitorRichTextBox.Clear();
                }
                _MonitorRichTextBox.AppendText(message + Environment.NewLine);
                _MonitorRichTextBox.ScrollToCaret();
            }
        }

        private void DisplayMessage(string message, System.Drawing.Color fontColor)
        {
            if (_MonitorRichTextBox.InvokeRequired)
            {
                _MonitorRichTextBox.Invoke(new MethodInvoker(() =>
                {
                    if (message.Length + Environment.NewLine.Length + _MonitorRichTextBox.TextLength >= 4194304)
                    {
                        _MonitorRichTextBox.Clear();
                    }
                    _MonitorRichTextBox.SelectionColor = fontColor;
                    _MonitorRichTextBox.AppendText(message + Environment.NewLine);
                    _MonitorRichTextBox.ScrollToCaret();
                }));
            }
            else
            {
                if (message.Length + Environment.NewLine.Length + _MonitorRichTextBox.TextLength >= 4194304)
                {
                    _MonitorRichTextBox.Clear();
                }
                _MonitorRichTextBox.AppendText(message + Environment.NewLine);
                _MonitorRichTextBox.ScrollToCaret();
            }
        }
        #endregion

        #region 高级文本框
        internal class MonitorRichTextBox : System.Windows.Forms.RichTextBox
        {
            protected override void OnPaintBackground(System.Windows.Forms.PaintEventArgs pevent)
            {
                return;
            }

            //protected override void OnPaint(System.Windows.Forms.PaintEventArgs e)
            //{
            //    e.Graphics.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.HighQuality;
            //    base.OnPaint(e);
            //}
        }
        #endregion
    }
}
