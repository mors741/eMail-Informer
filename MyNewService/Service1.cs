using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Linq;
using System.ServiceProcess;
using System.Text;
using System.Threading.Tasks;
using System.Net;
using System.Net.Mail;
using System.Net.Mime;
using System.Threading;

namespace MyNewService
{
    public partial class MyNewService : ServiceBase
    {
        private System.Timers.Timer timer;
        DataSet dataSet;
        System.Net.Mail.SmtpClient Smtp;

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

            Smtp = new SmtpClient("smtp.yandex.ru", 25);
            Smtp.EnableSsl = true;
            Smtp.Credentials = new System.Net.NetworkCredential("pot3r@yandex.ru", "plET56dfs0");
        }

        protected override void OnStart(string[] args)
        {
            eventLog1.WriteEntry("In OnStart");
            this.timer = new System.Timers.Timer(10000D);  // 30000 milliseconds = 30 seconds
            this.timer.AutoReset = true;
            this.timer.Elapsed += new System.Timers.ElapsedEventHandler(this.TimerElapsed);
            this.timer.Start();
            InitializeDataSet();
        }

        private void TimerElapsed(object sender, System.Timers.ElapsedEventArgs e)
        {
            eventLog1.WriteEntry("Timer Elapsed");
            SendMail(dataSet);
        }

        protected override void OnStop()
        {
            eventLog1.WriteEntry("In onStop.");
            this.timer.Stop();
            this.timer = null;
        }

        private void SendMail(DataSet dataSet)
        {
            { 
                int rowCount = dataSet.Tables[0].Rows.Count;
                for (int i=0; i<rowCount; i++){
                    MailMessage Message = new MailMessage();
                    Message.From = new MailAddress("pot3r@yandex.ru", "TMS-informer");
                    Message.To.Add(new MailAddress((String)dataSet.Tables[0].Rows[i][0]));
                    Message.Subject = (String) dataSet.Tables[0].Rows[i][1];
                    Message.Body = (String)dataSet.Tables[0].Rows[i][2];
                    // Message.Headers["X-mailer"] = "Headers";                   
                    Smtp.Send(Message);
                }
                
            }
        }

        private void SendMail(String body)  //depricated
        {
            {
                MailMessage Message = new MailMessage();
                Message.Subject = "Изменения в системе";
                Message.Body = body;
                // Message.Headers["X-mailer"] = "Headers"; 
                Message.From = new MailAddress("pot3r@yandex.ru", "TMS-informer");
                Message.To.Add(new MailAddress("mors741@yandex.ru"));
                System.Net.Mail.SmtpClient Smtp = new SmtpClient("smtp.yandex.ru", 25);
                //Smtp.DeliveryMethod = SmtpDeliveryMethod.Network;
                //Smtp.UseDefaultCredentials = false;
                Smtp.EnableSsl = true;
                Smtp.Credentials = new System.Net.NetworkCredential("pot3r@yandex.ru", "plET56dfs0");
                Smtp.Send(Message);
            }
        }

        private void PrintRows(DataSet dataSet)
        {
            // For each table in the DataSet, print the row values.
            foreach (DataTable table in dataSet.Tables)
            {
                foreach (DataRow row in table.Rows)
                {
                    foreach (DataColumn column in table.Columns)
                    {
                        Console.WriteLine(row[column]);
                    }
                }
            }
        }

        private void InitializeDataSet()
        {
            dataSet = new DataSet("DataSet");
            DataTable data = new DataTable("Data");
            DataColumn dtCol = null;

            dtCol = new DataColumn("Address");
            dtCol.DataType = typeof(System.String);
            data.Columns.Add(dtCol);

            dtCol = new DataColumn("Subject");
            dtCol.DataType = typeof(System.String);
            data.Columns.Add(dtCol);

            dtCol = new DataColumn("Text");
            dtCol.DataType = typeof(System.String);
            data.Columns.Add(dtCol);

            DataRow dr = null;
            dr = data.NewRow();
            dr["Address"] = "mors741@yandex.ru";
            dr["Subject"] = "Ошибка";
            dr["Text"] = "Терминал №89745019678 вышел из строя";
            data.Rows.Add(dr);

            dr = data.NewRow();
            dr["Address"] = "mors741@yandex.ru";
            dr["Subject"] = "Обновление";
            dr["Text"] = "Сегодяня вышло обновление системы версии 4.1.43";
            data.Rows.Add(dr);

            dr = data.NewRow();
            dr["Address"] = "mors741@yandex.ru";
            dr["Subject"] = "Устаревшее ПО";
            dr["Text"] = "На терминалах №89745019678-8974502004 требуется срочное обновление программного обеспечения!";
            data.Rows.Add(dr);

            dataSet.Tables.Add(data);
        }
        private void eventLog1_EntryWritten(object sender, EntryWrittenEventArgs e)
        {

        }
    }
}
