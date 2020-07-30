using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.Win32;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Management;
using System.Text;
using System.Windows.Forms;

namespace WFHTestApp
{
    static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Form1 formApp = new Form1();
            formApp.getCalendarData();
            formApp.UpdateMeetingStatus();
            formApp.UpdateCalendarStatus();
            Application.Run(formApp);
        }
        /// <summary>
        /// </summary>
        /// <param name="sender">
        /// The sender.
        /// </param>
        /// <param name="e">
        /// The e.
        /// </param>
        private static void HandleEvent(object sender, EventArrivedEventArgs e)
        {
            Console.WriteLine("Received an event.");
        }

    }
}
