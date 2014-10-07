using System;
using System.Data;
using System.Diagnostics;
using System.ServiceProcess;
using System.Net.Mail;

namespace MyNewService
{
    public partial class MyNewService : ServiceBase
    {
        private System.Timers.Timer timer;
        Informer informer;                                                                  //1

        public MyNewService()
        {
            InitializeComponent();
            if (!System.Diagnostics.EventLog.SourceExists("MySource"))
            {
                System.Diagnostics.EventLog.CreateEventSource(
                    "MySource", "MyNewLog");
            }
            eventLog1.Source = "MySource";
            eventLog1.Log = "MyNewLog";

            informer = new Informer();                                                       //2
        }

        protected override void OnStart(string[] args)
        {
            eventLog1.WriteEntry("In OnStart");
            this.timer = new System.Timers.Timer(10000D);  // 10000 milliseconds = 10 seconds
            this.timer.AutoReset = true;
            this.timer.Elapsed += new System.Timers.ElapsedEventHandler(this.TimerElapsed);
            this.timer.Start();
        }

        private void TimerElapsed(object sender, System.Timers.ElapsedEventArgs e)
        {
            eventLog1.WriteEntry("Timer Elapsed");
            informer.UpdateDataSet(/* */);                                                   //3
            informer.SendMail();                                                             //4
        }

        protected override void OnStop()
        {
            eventLog1.WriteEntry("In onStop.");
            this.timer.Stop();
            this.timer = null;
        }

        private void eventLog1_EntryWritten(object sender, EntryWrittenEventArgs e)
        {

        }
    }
}
