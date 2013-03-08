using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Management;
using System.Text;
using System.Threading;
using System.Windows.Forms;
using UplayStarter.Properties;
using Timer = System.Windows.Forms.Timer;

namespace UplayStarter
{
    public partial class UplayStarterForm : Form
    {
        private Timer _timer = new Timer();

        private ContextMenu _trayMenu;
        private NotifyIcon _trayIcon;
        private MenuItem _toggleItem;

        Process uPlayProcess = null;
        IntPtr uPlayWinHandle = IntPtr.Zero;
        int uPlayProcessId = int.MinValue;

        bool gameStartup = false;

        public UplayStarterForm()
        {
            InitializeComponent();

            MinimumSize = MaximumSize = Size;
            FormBorderStyle = FormBorderStyle.Fixed3D;

            _trayMenu = new ContextMenu();
            _toggleItem = new MenuItem("Enabled", (x,y) => checkBox1.Checked = !checkBox1.Checked);
            _toggleItem.Checked = checkBox1.Checked;
            _trayMenu.MenuItems.Add(_toggleItem);
            _trayMenu.MenuItems.Add(new MenuItem("Exit", (o, e) => Close()));

            _defaultBackColor = BackColor;
            UpdateAero();

            _trayIcon = new NotifyIcon();
            _trayIcon.Text = "Uplay Starter";
            _trayIcon.ContextMenu = _trayMenu;
            _trayIcon.Icon = Icon;
            _trayIcon.Visible = false;

            Closing += (o, e) =>
                           {
                               if (_trayIcon.Visible)
                               {
                                   _trayIcon.Visible = false;
                                   return;
                               }
                               _trayIcon.Visible = true;
                               _trayIcon.ShowBalloonTip(3000, "Uplay Starter Still Running!", "Uplay Starter will run in the background until closed.", ToolTipIcon.Info);
                               Visible = ShowInTaskbar = false;
                               e.Cancel = true;
                           };

            _trayIcon.MouseDoubleClick += (o, e) =>
                                              {
                                                  _trayIcon.Visible = false;
                                                  Visible = ShowInTaskbar = true;
                                              };

            VisibleChanged += (o, e) => UpdateAero();

            _timer.Interval = 500;
            _timer.Tick += (o, e) => processTick();
        }

        private Color _defaultBackColor;

        private void UpdateAero()
        {
            int en = 0;
            Win32API.MARGINS mg = new Win32API.MARGINS();
            mg.cxLeftWidth = mg.cxRightWidth = mg.cyTopHeight = mg.cyBottomHeight = checkBox1.Left * 2;

            bool transpaerent = false;
            //make sure you are not on a legacy OS 
            if (Environment.OSVersion.Version.Major >= 6)
            {
                Win32API.DwmIsCompositionEnabled(ref en);
                //check if the desktop composition is enabled

                if (en > 0)
                {
                    Win32API.DwmExtendFrameIntoClientArea(this.Handle, ref mg);
                    transpaerent = true;
                }
            }

            if (transpaerent)
                BackColor = Color.Black;
            else
                BackColor = DefaultBackColor;
        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            _toggleItem.Checked = checkBox1.Checked;
            if(checkBox1.Checked)
            {
                checkBox1.Text = Resources.Form1_checkBox1_CheckedChanged_Uplay_Starter_On;
                _timer.Enabled = true;
            }
            else
            {
                checkBox1.Text = Resources.Form1_checkBox1_CheckedChanged_Uplay_Starter_Off;
                _timer.Enabled = false;
            }
        }

        public static List<Tuple<string,int>> ListChildProcesses(int processID)
        {
            String machineName = "localhost";
            String myQuery = string.Format("select * from win32_process where ParentProcessId={0}", processID);
            ManagementScope mScope = new ManagementScope(string.Format(@"\\{0}\root\cimv2", machineName), null);
            mScope.Connect();

            var children = new List<Tuple<string, int>>();

            if (mScope.IsConnected)
            {
                ObjectQuery objQuery = new ObjectQuery(myQuery);
                using (ManagementObjectSearcher objSearcher = new ManagementObjectSearcher(mScope, objQuery))
                {
                    using (ManagementObjectCollection result = objSearcher.Get())
                    {
                        foreach (ManagementObject item in result)
                        {
                            children.Add(new Tuple<string, int>(item["Name"].ToString(), int.Parse(item["ProcessId"].ToString())));
                        }
                    }
                }
            }
            return children;
        }

        private void processTick()
        {
            if (uPlayProcess == null || uPlayProcessId == int.MinValue)
            {
                // Get uPlay process and window
                if (uPlayProcess == null)
                {
                    Console.WriteLine("Searching for uPlay");
                    foreach (Process proc in Process.GetProcesses())
                    {
                        if (proc.ProcessName.ToLower().Equals("uplay"))
                        {
                            Console.WriteLine("uPlay process found");
                            uPlayProcess = proc;
                            break;
                        }
                    }

                    if (uPlayProcess == null) return;

                    while (uPlayProcess.MainWindowHandle == IntPtr.Zero)
                    {
                        Console.WriteLine("Waiting for main window handle");
                        Thread.Sleep(500);
                    }

                    Console.WriteLine("Setting window handle");
                    uPlayWinHandle = uPlayProcess.MainWindowHandle;
                }

                uPlayProcessId = uPlayProcess.Id;
                Thread.Sleep(10000);
            }
            else if (!gameStartup)
            {
                // uPlay running, we need to wait for it to start up
                // then we send our keys
                Console.WriteLine("Sending input to uPlay");
                Win32API.SetForegroundWindow(uPlayWinHandle);

                SendKeys.SendWait("{TAB}");
                SendKeys.SendWait("{TAB}");
                SendKeys.SendWait("{ENTER}");
                SendKeys.Flush();

                // Then we wait for child processes to start up
                gameStartup = true;
                Thread.Sleep(5000);
            }
            else
            {
                List<Tuple<string, int>> children = ListChildProcesses(uPlayProcessId);

                if (children.Count > 0)
                {
                    foreach (Tuple<string, int> child in children)
                    {
                        Process proc = Process.GetProcessById(child.Item2);

                        if (proc != null)
                        {
                            proc.EnableRaisingEvents = true;
                            proc.Exited += childProcessExited;
                        }
                    }

                    _timer.Enabled = false;
                    gameStartup = false;
                }
            }
        }

        void childProcessExited(object sender, EventArgs e)
        {
            if (sender is Process)
            {
                Process senderProc = sender as Process;

                Thread.Sleep(1000);

                List<Tuple<string, int>> children = ListChildProcesses(uPlayProcessId);

                if (children.Count > 0)
                {
                    foreach (Tuple<string, int> child in children)
                    {
                        Process proc = Process.GetProcessById(child.Item2);

                        if (proc != null && proc != senderProc)
                        {
                            proc.EnableRaisingEvents = true;
                            proc.Exited += childProcessExited;
                        }
                    }
                }
                else
                {
                    CloseUPlay();
                }
            }
        }

        private void CloseUPlay()
        {
            Console.WriteLine("Killing uPlay process");

            Thread.Sleep(1000);
            uPlayProcess.Kill();
            _timer.Enabled = _toggleItem.Checked;
        }
    }
}
