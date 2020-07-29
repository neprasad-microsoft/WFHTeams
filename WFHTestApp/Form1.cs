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


namespace WFHTestApp
{
    public partial class Form1 : Form
    {
        bool ifVideoON = true;

        public Form1()
        {
            InitializeComponent();


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

                // Do something while waiting for events. In your application,
                // this would just be continuing business as usual.
         //       System.Threading.Thread.Sleep(100000000);

                // Stop listening for events.
            //    watcher.Stop();
            }
            catch (ManagementException managementException)
            {
                Console.WriteLine("An error occurred: " + managementException.Message);
            }





            ///Mocking an event with timer. CHange appropriately
            ///


            /*var timer = new Timer();
            timer.Interval = 2000;
            timer.Start();
            timer.Tick += new EventHandler(HandleEvent);                     
            // Stop listening for events.
            timer.Stop();*/
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
