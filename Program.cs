using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using System.Diagnostics;

namespace PBSepartor
{
    static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
        {

            // اجرا شدن یکبار برنامه
            string CurrentProcessName = Process.GetCurrentProcess().ProcessName;
            Process[] Processes = Process.GetProcessesByName(CurrentProcessName);
            if (Processes.Length > 1)
            {
                MessageBox.Show("this program is Running", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new Form1());
        }
    }
}
