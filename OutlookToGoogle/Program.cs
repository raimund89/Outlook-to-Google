using System;
using System.Collections.Generic;
using System.Security.Permissions;
using System.Windows.Forms;

// TODO: On selecting a folder, do some checks if 1) the file exists, 2) you have writing rights, 3) warn user in these cases
// TODO: On OK, do checks like 1) if file already exists, 2) sure to not start with the system, 3) etc.
// TODO: clicking on any item of the NotifyIcon fires the 'Click' action. Shouldn't happen.
// TODO: Program doesn't clean up nicely, RAM keeps increasing 2-3 MB every update. Wrong wrong wrong.
// TODO: File permissions can change. So check when setting a filename, on startup and on WriteICS
// TODO: Enter Google account details to immediately update the calendar there instead of using ICS

namespace OutlookToGoogle
{
    static class Program
    {
        public static OutlookICS ics = new OutlookICS();
        public static System.Threading.Timer updateTimer;
        public static NotifyIcon trayIcon;

        public static Dictionary<int, string> Intervals = new Dictionary<int, string>
        {
            { 0, "Every 5 minutes" },
            { 1, "Every 10 minutes" },
            { 2, "Every 30 minutes" },
            { 3, "Every 1 hour" },
            { 4, "Every 6 hours" },
            { 5, "Every 1 day" }
        };

        public static Dictionary<int, int> msIntervals = new Dictionary<int, int>
        {
            { 0,   300000 },
            { 1,   600000 },
            { 2,  1800000 },
            { 3,  3600000 },
            { 4, 21600000 },
            { 5, 86400000 }
        };

        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
        {
            /**************************/
            /*     Startup Checks     */
            /**************************/

            // Check if 'On system startup' should be set.
            // AKA, just set it according to the property
            ToggleStartup(Properties.Settings.Default.startWithSystem);

            // Check if the filename specified can be written to
            FileIOPermission fileIOPermission = new FileIOPermission(FileIOPermissionAccess.Write, GetICSPath());

            try
            {
                fileIOPermission.Demand();
            }
            catch(Exception e)
            {
                Console.WriteLine("Couldn't get permission to write the file: " + e.Message);
            }

            /**************************/
            /*    Application Entry   */
            /**************************/

            // Start the timer
            InitializeTimer();
            // And call it once for good measure

            // Start the tray icon
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new MyCustomApplicationContext());
        }

        public static void InitializeTimer()
        {
            if (updateTimer == null)
                updateTimer = new System.Threading.Timer(OnTimerFired, null, 0, msIntervals[Properties.Settings.Default.updateFreq]);
            else
                updateTimer.Change(0, msIntervals[Properties.Settings.Default.updateFreq]);
        }

        public static void OnTimerFired(Object stateInfo)
        {
            Console.WriteLine("Timer fired");
            ics.ReadCalendar();
            ics.WriteICS(GetICSPath());
            ics.Cleanup();

            if(Properties.Settings.Default.notifyOnChange)
                Program.trayIcon.ShowBalloonTip(1000, "OutlookToGoogle", "Calendar updated", ToolTipIcon.Info);
        }

        public static void ToggleStartup(bool startup)
        {
            if(startup)
            {
                Microsoft.Win32.RegistryKey key = Microsoft.Win32.Registry.CurrentUser.OpenSubKey("SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Run", true);
                key.SetValue("OutlookToGoogle", Application.ExecutablePath);
            } 
            else
            {
                Microsoft.Win32.RegistryKey key = Microsoft.Win32.Registry.CurrentUser.OpenSubKey("SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Run", true);
                key.DeleteValue("OutlookToGoogle", false);
            }
        }

        public static String GetICSPath()
        {
            return Properties.Settings.Default.icsPath + "\\" + Properties.Settings.Default.icsName + ".ics";
        }
    }

    public class MyCustomApplicationContext : ApplicationContext
    {
        private Form1 form1 = new Form1();

        public MyCustomApplicationContext()
        {
            // Initialize Tray Icon
            Program.trayIcon = new NotifyIcon()
            {
                Icon = Properties.Resources.AppIcon,
                ContextMenu = new ContextMenu(
                    new MenuItem[] {
                        new MenuItem("Settings", Settings),
                        new MenuItem("Update now", Update),
                        new MenuItem("-"),
                        new MenuItem("Exit", Exit)
                    }
                ),
                Visible = true
            };

            //Program.trayIcon.Click += new EventHandler(this.Settings);
            //Program.trayIcon.MouseUp += new MouseEventHandler(delegate (object sender, MouseEventArgs e) { Program.trayIcon.ContextMenu.Show(); });
        }

        void Exit(object sender, EventArgs e)
        {
            // Hide tray icon, otherwise it will remain shown until user mouses over it
            Program.trayIcon.Visible = false;

            // Cancel the timer
            Program.updateTimer.Dispose();

            Application.Exit();
        }

        void Settings(object sender, EventArgs e)
        {
            if(form1.Visible)
            {
                form1.Activate();
            }
            else
            {
                form1.ShowDialog();
            }
        }

        void Update(object sender, EventArgs e)
        {
            Program.OnTimerFired(null);
        }
    }
}
