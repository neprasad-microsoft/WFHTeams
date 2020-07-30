using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Management;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Timers;
using System.Net.Http;
using Newtonsoft.Json;



namespace WFHTestApp
{
    public partial class Form1 : Form
    {
        bool ifVideoON = true;
        private static System.Timers.Timer presenceTimer;
        private static System.Timers.Timer calendarTimer;

        private static List<DateTime> listOfTime = new List<DateTime>();
        private static DateTime nextMeetingTime;
        private static int nPresence;
        private static bool bMeetingReminder;


        public Form1()
        {
            InitializeComponent();

            getCalendarData();
            try
            {
                // Your query goes below; "KeyPath" is the key in the registry that you
                // want to monitor for changes. Make sure you escape the \ character.
                WqlEventQuery query = new WqlEventQuery(
                     "SELECT * FROM RegistryValueChangeEvent WHERE " +
                     "Hive = 'HKEY_LOCAL_MACHINE'" +
                     @"AND KeyPath = 'SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\CapabilityAccessManager\\ConsentStore\\webcam\\NonPackaged\\C:#Users#napulath#AppData#Local#Microsoft#Teams#current#Teams.exe' AND ValueName='LastUsedTimeStop'");

                ManagementEventWatcher watcher = new ManagementEventWatcher(query);
                //Console.WriteLine("Waiting for an event...");

                // Set up the delegate that will handle the change event.
                watcher.EventArrived += new EventArrivedEventHandler(HandleEvent);

                // Start listening for events.
                watcher.Start();

                // Stop listening for events.
            //    watcher.Stop();
            }
            catch (ManagementException managementException)
            {
                Console.WriteLine("An error occurred: " + managementException.Message);
            }


      
        }

        //Start timer for update meeting status
        public void UpdateMeetingStatus()
        {
            presenceTimer = new System.Timers.Timer();
            presenceTimer.Interval = 60000; //60 seconds
            presenceTimer.Elapsed += OnTimedEvent;
            presenceTimer.AutoReset = true;
            presenceTimer.Enabled = true;
            
        }
        public void UpdateCalendarStatus()
        {
            calendarTimer = new System.Timers.Timer();
            calendarTimer.Interval = 15 * 60000;
            calendarTimer.Elapsed += OnTimedEventCalendar;
            calendarTimer.AutoReset = true;
            calendarTimer.Enabled = true;

        }


        //Fetch and store calendar data
        private static void getCalendarData()
        {
            HttpClient client = new HttpClient();
            client.BaseAddress = new Uri("http://spring-boot-basewebapp-1595921330684.azurewebsites.net/");
             HttpResponseMessage response = client.GetAsync("getUserCalenderForNext24Hour").Result;  // Blocking call!

            if (response.IsSuccessStatusCode)
            {
                //var products = response.Content.ReadAsStringAsync().Result;
                var products = "{\"value\":[{\"start\":{\"dateTime\":\"2020 - 07 - 28T18: 00:00.0000000\",\"timeZone\":\"UTC\"},\"end\":{\"dateTime\":\"2020 - 07 - 28T19: 00:00.0000000\",\"timeZone\":\"UTC\"}},{\"start\":{\"dateTime\":\"2020 - 07 - 29T04: 30:00.0000000\",\"timeZone\":\"UTC\"},\"end\":{\"dateTime\":\"2020 - 07 - 29T06: 00:00.0000000\",\"timeZone\":\"UTC\"}},{\"start\":{\"dateTime\":\"2020 - 07 - 29T08: 30:00.0000000\",\"timeZone\":\"UTC\"},\"end\":{\"dateTime\":\"2020 - 07 - 29T10: 00:00.0000000\",\"timeZone\":\"UTC\"}},{\"start\":{\"dateTime\":\"2020 - 07 - 29T11: 00:00.0000000\",\"timeZone\":\"UTC\"},\"end\":{\"dateTime\":\"2020 - 07 - 29T12: 00:00.0000000\",\"timeZone\":\"UTC\"}}]}";
                var parser = Newtonsoft.Json.Linq.JObject.Parse(products);
                var value = parser["value"];//[1]["start"]["dateTime"];
                foreach (var item in value)
                {
                    //Console.WriteLine(item);
                    String str = item["start"]["dateTime"].ToObject<String>();
                    listOfTime.Add(DateTime.Parse(str).ToLocalTime());
                    //Console.WriteLine(DateTime.Parse(str));
                }
                
                listOfTime.Sort((a, b) => a.CompareTo(b));
                DateTime now = DateTime.UtcNow.ToLocalTime();
                DateTime datetime = listOfTime.First();

                while (datetime < now && listOfTime.Any())
                {
                    listOfTime.RemoveAt(0);
                    //Console.WriteLine("Removed time" + datetime.ToString());
                    if (listOfTime.Any())
                    {
                        datetime = listOfTime.First();
                    }
                    else
                    {
                        break;
                    }

                }

            }
            else
            {
                //Console.WriteLine("{0} ({1})", (int)response.StatusCode, response.ReasonPhrase);
            }

        }


        private static void OnTimedEventCalendar(Object source, System.Timers.ElapsedEventArgs e)
        {
            DateTime nextMeeting = listOfTime.First();
            listOfTime.RemoveAt(0);
            nextMeetingTime = nextMeeting;
            DateTime now = DateTime.UtcNow.ToLocalTime();
            TimeSpan span = nextMeetingTime.Subtract(now);
            if (span.TotalMinutes.CompareTo(15d) < 0)
            {
                bMeetingReminder = true;
                //Console.WriteLine("Time Difference (minutes): " + span.TotalMinutes);
                refreshScreen();
            }
            else
            {
                bMeetingReminder = false;
            }
            String nextMeetingStr = nextMeeting.ToString("dd/MM/yyyy HH:mm:ss");
            //Console.WriteLine("The Next Meeting is at "+nextMeetingStr);
        }


        //GetUserPresence
        private static string getPresence()
        {
            HttpClient client = new HttpClient();
            client.BaseAddress = new Uri("http://spring-boot-basewebapp-1595921330684.azurewebsites.net/");
            HttpResponseMessage response = client.GetAsync("getUsersPresence").Result;  // Blocking call!
            if (response.IsSuccessStatusCode)
            {
                //var retVal = response.Content.ReadAsStringAsync().Result;
                var retVal = "{ \"@odata.context\":\"https://graph.microsoft.com/beta/$metadata#users('206b5a0b-3b77-4148-b5bb-77e0187c9347')/presence/$entity\",\"id\":\"206b5a0b-3b77-4148-b5bb-77e0187c9347\",\"availability\":\"Available\",\"activity\":\"Available\"}";
                var parser = Newtonsoft.Json.Linq.JObject.Parse(retVal);
                //Console.WriteLine(parser);
                var presence = parser["availability"];
                String strPresence =  presence.ToObject<String>();
                if (strPresence.Equals("Available", StringComparison.OrdinalIgnoreCase))
                {
                    nPresence = 1;
                } else
                {
                    nPresence = 0;
                }
               
            }
            else
            {
                //Console.WriteLine("{0} ({1})", (int)response.StatusCode, response.ReasonPhrase);
            }

            return null;
        }

        private static void refreshScreen()
        {
            //Use combination of fVideoOn, nPresence and bMeetingReminder to display different images
        }

        private static void OnTimedEvent(Object source, System.Timers.ElapsedEventArgs e)
        {

            //Console.WriteLine(getPresence());
            refreshScreen();
            //Console.WriteLine("The Elapsed event was raised at {0}", e.SignalTime);
        }
        /// <summary>
        /// </summary>
        /// <param name="sender">
        /// The sender.
        /// </param>
        /// <param name="e">
        /// The e.
        /// </param>
        /// 
        /// -----------change EventArgs to EventArrivedEventArgs
        private void HandleEvent(object sender, EventArgs e)
        {
        
            //Console.WriteLine("Received an event.");
            //opening the subkey  
            RegistryKey key = Registry.LocalMachine.OpenSubKey(@"SOFTWARE\Microsoft\Windows\CurrentVersion\CapabilityAccessManager\ConsentStore\webcam\NonPackaged\C:#Users#napulath#AppData#Local#Microsoft#Teams#current#Teams.exe");
            //if it does exist, retrieve the stored values  
            if (key != null)
            {
   
                Object obj = key.GetValue("LastUsedTimeStop");
                //int nVal = (int)(obj as int?);
                long nVal = (long)(obj);
                if (nVal == 0)
                {
                    ifVideoON = true;
                }
                else
                {
                    ifVideoON = false;
                }

                refreshScreen();
                key.Close();
            }

        }

        public void VideoON()
        {
//            pictureBox1.Image = Image.FromFile(@"C:\Users\napulath\source\repos\WFHTestApp\VideoOn.jpg"); // global::WFHTestApp.Properties.Resources.MicOff; 
        }

        public void VideoOFF()
        {
  //          pictureBox1.Image = Image.FromFile(@"C:\Users\napulath\source\repos\WFHTestApp\VideoOff.jpg");
        }

        public void MicOFF()
        {
            //pictureBox2.Image = Image.FromFile(@"C:\Users\napulath\source\repos\WFHTestApp\MicOff.jpg");
        }

        public void MicON()
        {
  //          pictureBox2.Image = Image.FromFile(@"C:\Users\napulath\source\repos\WFHTestApp\MicOn.jpg");
        }

        public void SpeakerOFF()
        {
  //          pictureBox3.Image = Image.FromFile(@"C:\Users\napulath\source\repos\WFHTestApp\SpeakerOff.jpg");
        }

        public void SpeakerON()
        {            
   //        pictureBox3.Image = Image.FromFile(@"C:\Users\napulath\source\repos\WFHTestApp\SpeakerOn.jpg");  
        }

        public void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Form1));
            this.pictureBox1 = new System.Windows.Forms.PictureBox();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).BeginInit();
            this.SuspendLayout();
            // 
            // pictureBox1
            // 
            this.pictureBox1.Image = ((System.Drawing.Image)(resources.GetObject("pictureBox1.Image")));
            this.pictureBox1.Location = new System.Drawing.Point(0, 2);
            this.pictureBox1.Name = "pictureBox1";
            this.pictureBox1.Size = new System.Drawing.Size(654, 430);
            this.pictureBox1.TabIndex = 0;
            this.pictureBox1.TabStop = false;
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(9F, 20F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(655, 432);
            this.Controls.Add(this.pictureBox1);
            this.Name = "Form1";
            this.Text = "Form1";
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).EndInit();
            this.ResumeLayout(false);

        }

    }
}
