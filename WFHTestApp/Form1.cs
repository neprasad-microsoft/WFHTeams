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
using System.Threading.Tasks;
using Newtonsoft.Json;



namespace WFHTestApp
{
    public partial class Form1 : Form
    {
        bool ifVideoON = true;
        private static System.Timers.Timer presenceTimer;
        private static System.Timers.Timer calendarTimer;

        private static List<DateTime> listOfTime = new List<DateTime>();


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
                     @"AND KeyPath = 'SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\CapabilityAccessManager\\ConsentStore\\webcam\\NonPackaged\\C:#Users#neprasad#AppData#Local#Microsoft#Teams#current#Teams.exe' AND ValueName='LastUsedTimeStop'");

                ManagementEventWatcher watcher = new ManagementEventWatcher(query);
                Console.WriteLine("Waiting for an event...");

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
                    Console.WriteLine(item);
                    String str = item["start"]["dateTime"].ToObject<String>();
                    listOfTime.Add(DateTime.Parse(str).ToLocalTime());
                    Console.WriteLine(DateTime.Parse(str));
                }
                
                listOfTime.Sort((a, b) => a.CompareTo(b));
                DateTime now = DateTime.Now.ToLocalTime();
                DateTime datetime = listOfTime.First();

                while (datetime < now && listOfTime.Any())
                {
                    listOfTime.RemoveAt(0);
                    Console.WriteLine("Removed time" + datetime.ToString());
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
                Console.WriteLine("{0} ({1})", (int)response.StatusCode, response.ReasonPhrase);
            }

        }


        private static void OnTimedEventCalendar(Object source, System.Timers.ElapsedEventArgs e)
        {
            DateTime nextMeeting = listOfTime.First();
            listOfTime.RemoveAt(0);
            String nextMeetingStr = nextMeeting.ToString("dd/MM/yyyy HH:mm:ss");
            Console.WriteLine("The Next Meeting is at "+nextMeetingStr);
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
                Console.WriteLine(parser);
                var presence = parser["availability"];
                return presence.ToObject<String>();

            }
            else
            {
                Console.WriteLine("{0} ({1})", (int)response.StatusCode, response.ReasonPhrase);
            }

            return null;
        }

        private static void OnTimedEvent(Object source, System.Timers.ElapsedEventArgs e)
        {

            Console.WriteLine(getPresence());
            Console.WriteLine("The Elapsed event was raised at {0}", e.SignalTime);
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
        
            Console.WriteLine("Received an event.");
            //opening the subkey  
            RegistryKey key = Registry.LocalMachine.OpenSubKey(@"SOFTWARE\Microsoft\Windows\CurrentVersion\CapabilityAccessManager\ConsentStore\webcam\NonPackaged\C:#Users#neprasad#AppData#Local#Microsoft#Teams#current#Teams.exe");
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
                if (ifVideoON)
                {
                    VideoON();

                }
                else
                {
                    VideoOFF();
                }
                key.Close();
            }

        }

        private void pictureBox1_Click(object sender, EventArgs e)
        {

        }

        private void pictureBox2_Click(object sender, EventArgs e)
        {

        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        public void VideoON()
        {
            pictureBox1.Image = Image.FromFile(@"C:\Users\neprasad\source\repos\WFHTestApp\VideoOn.jpg"); // global::WFHTestApp.Properties.Resources.MicOff; 
        }

        public void VideoOFF()
        {
            pictureBox1.Image = Image.FromFile(@"C:\Users\neprasad\source\repos\WFHTestApp\VideoOff.jpg");
        }

        public void MicOFF()
        {
            pictureBox2.Image = Image.FromFile(@"C:\Users\neprasad\source\repos\WFHTestApp\MicOff.jpg");
        }

        public void MicON()
        {
            pictureBox2.Image = Image.FromFile(@"C:\Users\neprasad\source\repos\WFHTestApp\MicOn.jpg");
        }

        public void SpeakerOFF()
        {
            pictureBox3.Image = Image.FromFile(@"C:\Users\neprasad\source\repos\WFHTestApp\SpeakerOff.jpg");
        }

        public void SpeakerON()
        {            
            pictureBox3.Image = Image.FromFile(@"C:\Users\neprasad\source\repos\WFHTestApp\SpeakerOn.jpg");  
        }


        public void InitializeComponent()
        {
            this.pictureBox3 = new System.Windows.Forms.PictureBox();
            this.pictureBox2 = new System.Windows.Forms.PictureBox();
            this.pictureBox1 = new System.Windows.Forms.PictureBox();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox3)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox2)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).BeginInit();
            this.SuspendLayout();


            // 
            // pictureBox3
            // 
            this.pictureBox3.Image = Image.FromFile(@"C:\Users\neprasad\source\repos\WFHTestApp\SpeakerOff.jpg");
            this.pictureBox3.Location = new System.Drawing.Point(443, 117);
            this.pictureBox3.Name = "pictureBox3";
            this.pictureBox3.Size = new System.Drawing.Size(131, 130);
            this.pictureBox3.TabIndex = 2;
            this.pictureBox3.TabStop = false;
            // 
            // pictureBox2
            // 
            this.pictureBox2.Image = Image.FromFile(@"C:\Users\neprasad\source\repos\WFHTestApp\MicOff.jpg");
            this.pictureBox2.Location = new System.Drawing.Point(240, 117);
            this.pictureBox2.Name = "pictureBox2";
            this.pictureBox2.Size = new System.Drawing.Size(134, 130);
            this.pictureBox2.TabIndex = 1;
            this.pictureBox2.TabStop = false;
            this.pictureBox2.Click += new System.EventHandler(this.pictureBox2_Click);

            // 
            // pictureBox1
            // 
            this.pictureBox1.Image = Image.FromFile(@"C:\Users\neprasad\source\repos\WFHTestApp\VideoOff.jpg");
            this.pictureBox1.Location = new System.Drawing.Point(50, 117);
            this.pictureBox1.Name = "pictureBox1";
            this.pictureBox1.Size = new System.Drawing.Size(124, 130);
            this.pictureBox1.TabIndex = 0;
            this.pictureBox1.TabStop = false;
            this.pictureBox1.Click += new System.EventHandler(this.pictureBox1_Click);            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(9F, 20F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(639, 308);
            this.Controls.Add(this.pictureBox3);
            this.Controls.Add(this.pictureBox2);
            this.Controls.Add(this.pictureBox1);
            this.Name = "Form1";
            this.Text = "Form1";
            this.Load += new System.EventHandler(this.Form1_Load);
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox3)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox2)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).EndInit();
            this.ResumeLayout(false);
        }
            
             public System.Windows.Forms.PictureBox pictureBox1;
             public System.Windows.Forms.PictureBox pictureBox2;
             public System.Windows.Forms.PictureBox pictureBox3; 
    }
}
