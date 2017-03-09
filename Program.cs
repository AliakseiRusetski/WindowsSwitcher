using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Windows.Forms;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Text;
using System.ServiceProcess;
using System.ComponentModel;
using System.Configuration.Install;
using Microsoft.Win32;
using System.Linq;

namespace ConsoleApplication
{
    /* [RunInstaller(true)]
     public class WindowsServiceInstaller : Installer
     {
         /// <summary>
         /// Public Constructor for WindowsServiceInstaller.
         /// - Put all of your Initialization code here.
         /// </summary>
         public WindowsServiceInstaller()
         {
             ServiceProcessInstaller serviceProcessInstaller =
                                new ServiceProcessInstaller();
             ServiceInstaller serviceInstaller = new ServiceInstaller();

             //# Service Account Information
             serviceProcessInstaller.Account = ServiceAccount.LocalSystem;
             serviceProcessInstaller.Username = null;
             serviceProcessInstaller.Password = null;

             //# Service Information
             serviceInstaller.DisplayName = "My New C# Windows Service";
             serviceInstaller.StartType = ServiceStartMode.Automatic;

             //# This must be identical to the WindowsService.ServiceBase name
             //# set in the constructor of WindowsService.cs
             serviceInstaller.ServiceName = "My Windows Service";

             this.Installers.Add(serviceProcessInstaller);
             this.Installers.Add(serviceInstaller);
         }
     }*/

    public class Program : ServiceBase
    {

        private const int WH_KEYBOARD_LL = 13;
        private const int WM_KEYDOWN = 0x0100;
        private static LowLevelKeyboardProc keyboardProc = HookCallback;
        private static IntPtr keyboardHook = IntPtr.Zero;
        private static int[] lastPressedKeys = new int[] { 0, 0 };
        private static DateTime? lastPressTime;
        private static int[] VK_TEMPLATE = new int[] { 162, 91 }; // ctrl, win
        private static Program main;

        public static void Main(string[] args)
        {
            // ServiceBase.Run(new Program());
            //ShowForm();

            main = new Program();
            Application.EnableVisualStyles();
            ShowForm();
            Application.Run(selector);
            main.OnStop();

        }

        Program()
        {
            this.ServiceName = "SwitcherHook";
            this.EventLog.Log = "Launched";
            this.CanStop = true;
            keyboardHook = SetHook(keyboardProc);
        }

        protected override void OnStop()
        {
            UnhookWindowsHookEx(keyboardHook);
            // base.OnStop();
        }

        private static Form selector;
        private static void ShowForm()
        {
            Console.WriteLine("ShowForm");
            if (selector == null || selector.IsDisposed)
            {
                selector = new WindowsSelector();
            }
            if (selector != null)
            {
                selector.Show();
                Console.WriteLine("Show window");
                ShowAppWindow(selector.Handle, false);
                selector.Activate();
            }
        }

        internal static void ShowAppWindow(IntPtr hWnd, bool useTabMethod = true)
        {
            IntPtr attachTo = GetFocus();
            IntPtr switcher = GetCurrentThreadId();
            AttachThreadInput(switcher, attachTo, true);
            Console.WriteLine("Attaching: " + switcher.ToString() + " to " + attachTo.ToString() + " useTab:" + useTabMethod);
            BringWindowToTop(hWnd);
            if (useTabMethod)
            {
                SwitchToThisWindow(hWnd, true);
            }
            else
            {
                SetForegroundWindow(hWnd);
            }

            Console.WriteLine("Dettaching: " + switcher.ToString() + " from " + attachTo.ToString());
            AttachThreadInput(switcher, attachTo, false);
        }

        private static IntPtr SetHook(LowLevelKeyboardProc proc)
        {
            using (Process curProcess = Process.GetCurrentProcess())
            using (ProcessModule curModule = curProcess.MainModule)
            {
                return SetWindowsHookEx(WH_KEYBOARD_LL, proc,
                    GetModuleHandle(curModule.ModuleName), 0);
            }
        }

        private delegate IntPtr LowLevelKeyboardProc(int nCode, IntPtr wParam, IntPtr lParam);

        private static IntPtr HookCallback(int nCode, IntPtr wParam, IntPtr lParam)
        {
            if (nCode >= 0 && wParam == (IntPtr)WM_KEYDOWN)
            {
                int vkCode = Marshal.ReadInt32(lParam);
                PutKey(vkCode);
                if (IsKeysMatch())
                {
                    ShowForm();
                }
            }

            return CallNextHookEx(keyboardHook, nCode, wParam, lParam);
        }

        private static bool IsKeysMatch()
        {
            return VK_TEMPLATE.SequenceEqual(lastPressedKeys);
        }

        private static void PutKey(int vkCode)
        {
            if (lastPressTime != null && DateTime.Now.AddSeconds(-2) > lastPressTime)
            {
                lastPressTime = null;
                Array.Clear(lastPressedKeys, 0, lastPressedKeys.Length);
            }
            Array.Copy(lastPressedKeys, 1, lastPressedKeys, 0, lastPressedKeys.Length - 1);
            lastPressedKeys[lastPressedKeys.Length - 1] = vkCode;
            lastPressTime = DateTime.Now;
        }

        [DllImport("kernel32.dll")]
        internal static extern IntPtr GetConsoleWindow();

        [DllImport("user32.dll", CharSet = CharSet.Auto)]
        private static extern IntPtr SetWindowsHookEx(int idHook, LowLevelKeyboardProc lpfn, IntPtr hMod, uint dwThreadId);

        [DllImport("user32.dll", CharSet = CharSet.Auto)]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool UnhookWindowsHookEx(IntPtr hhk);

        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        internal static extern bool SetForegroundWindow(IntPtr hWnd);

        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        internal static extern bool AttachThreadInput(IntPtr idAttach, IntPtr idAttachTo, bool fAttach);

        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        internal static extern bool IsWindowVisible(IntPtr hWnd);

        [DllImport("user32.dll")]
        internal static extern IntPtr SetActiveWindow(IntPtr hWnd);

        [DllImport("user32.dll")]
        internal static extern IntPtr BringWindowToTop(IntPtr hWnd);

        [DllImport("user32.dll", CharSet = CharSet.Auto)]
        private static extern IntPtr CallNextHookEx(IntPtr hhk, int nCode, IntPtr wParam, IntPtr lParam);

        [DllImport("kernel32.dll", CharSet = CharSet.Auto)]
        private static extern IntPtr GetModuleHandle(string lpModuleName);

        [DllImport("user32.dll")]
        internal static extern bool ShowWindow(IntPtr hwnd, int nCmdShow);

        [DllImport("user32.dll")]
        internal static extern IntPtr GetForegroundWindow();

        [DllImport("user32.dll", SetLastError = true)]
        internal static extern IntPtr SetFocus(IntPtr hWnd);

        [DllImport("user32.dll")]
        internal static extern IntPtr GetFocus();

        [DllImport("kernel32.dll")]
        internal static extern IntPtr GetCurrentThreadId();

        [DllImport("user32.dll")]
        internal static extern void SwitchToThisWindow(IntPtr hWnd, bool turnOn);

        [DllImport("user32.dll", CharSet = CharSet.Auto)]
        private static extern int GetWindowText(IntPtr hWnd, StringBuilder strText, int maxCount);

        [DllImport("user32.dll", CharSet = CharSet.Auto)]
        private static extern int GetWindowTextLength(IntPtr hWnd);

        [DllImport("user32.dll")]
        private static extern bool EnumWindows(EnumWindowsProc enumProc, IntPtr lParam);
        [DllImport("user32.dll", SetLastError = true)]
        internal static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint processId);

        public static IEnumerable<IntPtr> FindWindows(EnumWindowsProc filter)
        {
            IntPtr found = IntPtr.Zero;
            List<IntPtr> windows = new List<IntPtr>();

            EnumWindows(delegate (IntPtr wnd, IntPtr param)
            {
                if (filter(wnd, param))
                {
                    windows.Add(wnd);
                }
                return true;
            }, IntPtr.Zero);

            return windows;
        }
        public static string GetWindowText(IntPtr hWnd)
        {
            int size = GetWindowTextLength(hWnd);
            if (size > 0)
            {
                var builder = new StringBuilder(size + 1);
                GetWindowText(hWnd, builder, builder.Capacity);
                return builder.ToString();
            }

            return String.Empty;
        }
    }

    // Delegate to filter which windows to include 
    public delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);

    class WindowsSelector : Form
    {
        private const string AUTORUN_KEY = "WindowsSwitcher";
        private TextBox txt;
        private CheckBox autorunCheckbox;
        private ListBox list;
        private TableLayoutPanel pan;
        private List<ProcessInfo> processInfos;

        public WindowsSelector()
        {
            Text = "WindowsSelector";
            Shown += WindowsSelector_ShownFirst;
            Shown += WindowsSelector_Shown;
            VisibleChanged += WindowsSelector_Shown;
            FormBorderStyle = FormBorderStyle.FixedSingle;
            Size = new Size(650, 800);
            TopMost = true;
            InitComponents();
        }

        private void WindowsSelector_ShownFirst(Object sender, EventArgs e)
        {
            Hide();
            IntPtr con = ConsoleApplication.Program.GetConsoleWindow();
            if (IntPtr.Zero != con)
            {
                ConsoleApplication.Program.ShowWindow(con, 0);
            }
        }
        private void WindowsSelector_Shown(Object sender, EventArgs e)
        {
            if (!Visible) return;
            txt.Clear();
            GetAltTabWindows();
            DisplayWindows();
            ActiveControl = txt;

            Console.WriteLine("onShown");
        }

        private void DisplayWindows()
        {
            list.Items.Clear();
            processInfos.ForEach(proc => list.Items.Add(proc.ToString()));
            list.SelectedIndex = 0;
            list.SelectionMode = SelectionMode.One;
        }

        private void InitComponents()
        {
            pan = InitPan();
            txt = new TextBox();
            txt.Multiline = false;
            txt.ShortcutsEnabled = true;
            list = new ListBox();
            autorunCheckbox = new CheckBox();
            int innerWidth = (int)(Size.Width * 0.97);
            int innerHeight = (int)(Size.Height * 0.97);
            txt.Width = innerWidth;
            txt.KeyDown += Selector_KeyDown;
            list.KeyDown += Selector_KeyDown;
            list.Width = innerWidth;
            list.Height = (int)(innerHeight * 0.90);
            autorunCheckbox.Text = "Autorun";
            autorunCheckbox.Checked = GetAutorunRK().GetValue(AUTORUN_KEY) != null;
            autorunCheckbox.CheckedChanged += Selector_OnAutorunCheckChanged;
            Controls.Add(pan);
            pan.Controls.Add(txt);
            pan.Controls.Add(list);
            pan.Controls.Add(autorunCheckbox);
            txt.TextChanged += txt_TextChanged;
        }

        private RegistryKey GetAutorunRK()
        {
            RegistryKey rk = Registry.CurrentUser.OpenSubKey("SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Run", true);
            return rk;
        }

        private void Selector_OnAutorunCheckChanged(object sender, EventArgs e)
        {
            var rk = GetAutorunRK();
            if (autorunCheckbox.Checked)
                rk.SetValue(AUTORUN_KEY, Application.ExecutablePath.ToString());
            else
                rk.DeleteValue(AUTORUN_KEY, false);
        }

        private void Selector_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                if (processInfos.Count > 0 && list.SelectedIndex >= 0 && list.SelectedIndex < processInfos.Count)
                {
                    IntPtr hWnd = processInfos[list.SelectedIndex].hwnd;
                    Console.WriteLine("showing: " + hWnd.ToString());
                    ConsoleApplication.Program.ShowAppWindow(hWnd);
                }
                e.Handled = true;
                e.SuppressKeyPress = true;
                // Close();
                Hide();
            }
            if (e.KeyCode == Keys.Escape)
            {
                // Close();
                e.Handled = true;
                e.SuppressKeyPress = true;
                Hide();
            }
        }

        private void txt_TextChanged(object sender, EventArgs e)
        {
            string txtValue = txt.Text;
            UpdateFilter(txtValue);
            DisplayWindows();
        }

        private void UpdateFilter(string txtValue)
        {
            processInfos.ForEach(ps => ps.Recalc(txtValue));
            processInfos.Sort();
        }

        private TableLayoutPanel InitPan()
        {
            TableLayoutPanel pan = new TableLayoutPanel();
            pan.Location = new Point(0, 0);
            pan.Size = Size;
            pan.AutoSize = true;
            pan.Name = "Desk";
            pan.ColumnCount = 0;
            pan.RowCount = 2;
            pan.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            pan.GrowStyle = System.Windows.Forms.TableLayoutPanelGrowStyle.AddRows;
            return pan;
        }

        private void GetAltTabWindows()
        {
            processInfos = new List<ProcessInfo>();
            IEnumerable<IntPtr> windows = ConsoleApplication.Program.FindWindows(delegate (IntPtr hWnd, IntPtr lParam)
               {
                   return true;
               });

            foreach (IntPtr window in windows)
            {
                uint pid;
                ConsoleApplication.Program.GetWindowThreadProcessId(window, out pid);
                if (!ConsoleApplication.Program.IsWindowVisible(window)) continue;
                string title = ConsoleApplication.Program.GetWindowText(window);
                if (string.IsNullOrEmpty(title)) continue;
                string name = Process.GetProcessById((int)pid).ProcessName;
                processInfos.Add(new ProcessInfo((int)pid, window, name, title));
            }
        }
    }

    class ProcessInfo : IComparable
    {

        public ProcessInfo(int pid, IntPtr window, string name, string title)
        {
            this.pid = pid;
            this.hwnd = window;
            this.name = name;
            this.title = title;
        }

        public int weight { get; private set; }
        public IntPtr hwnd { get; private set; }
        public int pid { get; private set; }
        public string name { get; private set; }
        public string title { get; private set; }

        override public string ToString()
        {
            return String.Format("pid:{0,-5}  hwnd:{1,-14} {2}>>>>{3}", pid, hwnd.ToInt64(), name, title);
        }

        internal void Recalc(string txtValue)
        {
            string lower = (txtValue.Clone() as string).ToLower();
            string nameLower = name.ToLower();
            string titleLower = title.ToLower();
            weight = int.MaxValue;
            if (nameLower.Contains(txtValue) || titleLower.Contains(txtValue))
            {
                weight = 1;
            }
            weight = new int[]{
                weight,
                LevenshteinDistance.Compute(lower, nameLower),
                LevenshteinDistance.Compute(lower, titleLower),
                LevenshteinDistance.Compute(lower, pid.ToString())
            }.Min();
        }

        public int CompareTo(object obj)
        {
            if (obj == null) return 1;

            ProcessInfo otherProcessInfo = obj as ProcessInfo;
            if (otherProcessInfo != null)
                return this.weight.CompareTo(otherProcessInfo.weight);
            else
                throw new ArgumentException("Object is not a ProcessInfo");
        }
    }

    static class LevenshteinDistance
    {
        public static int Compute(string s, string t)
        {
            if (string.IsNullOrEmpty(s))
            {
                if (string.IsNullOrEmpty(t))
                    return 0;
                return t.Length;
            }

            if (string.IsNullOrEmpty(t))
            {
                return s.Length;
            }

            int n = s.Length;
            int m = t.Length;
            int[,] d = new int[n + 1, m + 1];

            // initialize the top and right of the table to 0, 1, 2, ...
            for (int i = 0; i <= n; d[i, 0] = i++) ;
            for (int j = 1; j <= m; d[0, j] = j++) ;

            for (int i = 1; i <= n; i++)
            {
                for (int j = 1; j <= m; j++)
                {
                    int cost = (t[j - 1] == s[i - 1]) ? 0 : 1;
                    int min1 = d[i - 1, j] + 1;
                    int min2 = d[i, j - 1] + 1;
                    int min3 = d[i - 1, j - 1] + cost;
                    d[i, j] = Math.Min(Math.Min(min1, min2), min3);
                }
            }
            return d[n, m];
        }
    }
}
