using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Net.Mail;
using System.Text;
using System.Threading.Tasks;

namespace mailScheduler
{
    public class SendMail
    {
        public void SendEmail(string MailTo, string MailSubject)
        {
            try
            {
                string AppLocation = "";
                AppLocation = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().CodeBase);
                AppLocation = AppLocation.Replace("file:\\", "");
                string file = AppLocation + "\\ExcelFiles\\Bet9jaReport.xlsx";

                MailMessage mail = new MailMessage();
                SmtpClient SmtpServer = new SmtpClient("smtp.gmail.com", 587);
                DateTime previousDay = DateTime.Now.Date.AddDays(-1);

                mail.From = new MailAddress("skulkid51@gmail.com");
                mail.To.Add(MailTo);
                List<string> li = new List<string>();
                li.Add("whitehombre@outlook.com");
                li.Add("chiamakamushinwa@gmail.com");
                mail.CC.Add(string.Join(",", li));
                mail.Bcc.Add(string.Join(",", li));
                mail.Subject = MailSubject;
                mail.Body = "Hello Iretioluwa,\n" +
                    "I hope you are doing well.\n" +
                    "Please, find attached report for Bet9ja transactions done yesterday " + previousDay;

                System.Net.Mail.Attachment attachment;
                attachment = new System.Net.Mail.Attachment(file);
                mail.Attachments.Add(attachment);

                // Set SMTP server credentials and enable SSL/TLS if required
                SmtpServer.Credentials = new System.Net.NetworkCredential("skulkid51@gmail.com", "nbjccuqeixvsewde");
                SmtpServer.EnableSsl = true;

               SmtpServer.Send(mail);
               Console.WriteLine("Eamil sent successfully");
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
    }
}
