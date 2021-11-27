using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using System.ComponentModel;
using System.Net.Mail;


namespace IanWongHelpers
{
    public class SmtpHelper
    {
        private string _username;
        private string _password;

        private SmtpClient _client;

        /// <summary>
        /// 
        /// </summary>
        /// <param name="host"></param>
        /// <param name="port"></param>
        /// <param name="enableSSL"></param>
        /// <param name="username"></param>
        /// <param name="password"></param>
        /// <param name="timeout">timeout，默認120秒，單位毫秒</param>
        public SmtpHelper(string host, int port, bool enableSSL, string username, string password,int timeout=120000)
        {
            _username = username;
            _password = password;
            _client = new SmtpClient(host,port);
            _client.UseDefaultCredentials = false;
            _client.Credentials = new System.Net.NetworkCredential(username, password);
            _client.EnableSsl = enableSSL;
            _client.Timeout = timeout;  
            _client.SendCompleted += new SendCompletedEventHandler(SmtpClientSendCompleted); //註冊一個Callback，在每封email發送完成（包括成功失敗）時處理其他事項，如失敗時記錄Logs
        }

        public void SendAsync(MailMessage mail)
        {
            _client.SendAsync(mail, mail);
        }


        public void SendAsync(string subject, string body, bool bodyIsHtml, string from, string to, List<Attachment>? attachments )
        {

            SendAsync(subject, body, bodyIsHtml, from, new string[] { to }, null, null, attachments);
        }

        public void SendAsync(string subject, string body, bool bodyIsHtml, string from, string[] to, string[]? cc, string[]? bcc, List<Attachment>? attachments)
        {

            MailMessage mail = new MailMessage();

            mail.From = new MailAddress(from);//對於Sendgrid，sender必須是在Sendgrid中已創建。查看Sendgrid sender: Dashboard->Maketing->Senders

            foreach (var address in to)
            {
                mail.To.Add(address);
            }

            if (cc != null)
            {
                foreach (var address in cc)
                {
                    mail.CC.Add(address);
                }
            }

            if (bcc != null)
            {
                foreach (var address in bcc)
                {
                    mail.Bcc.Add(address);
                }
            }

            mail.IsBodyHtml = bodyIsHtml;
            mail.Subject = subject;
            mail.Body = body;

            if (attachments != null) 
            { 
                foreach (var attachment in attachments)
                {
                    mail.Attachments.Add(attachment);
                }
            }

            //將物理文件作為附件發送
            //var fileAttachment = new System.Net.Mail.Attachment("d:/textfile.txt");
            //mail.Attachments.Add(fileAttachment);

            //將Stream作為附件發送
            //FileStream fs = new FileStream("d:/textfile2.txt", FileMode.Open);
            //var streamAttachment = new Attachment(fs, "fileFromStream.txt");
            //mail.Attachments.Add(streamAttachment);

            _client.SendAsync(mail, mail);
        }

        void SmtpClientSendCompleted(object sender, AsyncCompletedEventArgs e)
        {
            var smtpClient = (SmtpClient)sender;
           
            smtpClient.SendCompleted -= SmtpClientSendCompleted;

            if (e.Error != null) //發送郵件失敗，發生錯誤
            {
                //Send email failed, do something here

                if (e.UserState != null)
                { 
                    var mailMessage = (MailMessage)e.UserState;
                    foreach (var address in mailMessage.To)
                    {
                        File.AppendAllLines("d:/test-log.txt", new string[] { $" send email to {address} failed." });
                        File.AppendAllLines("d:/test-log.txt", new string[] { $" failed reason: {e.Error.ToString()}" });
                        File.AppendAllLines("d:/test-log.txt", new string[] { $"\r\n \r\n \r\n" });
                    }
                }                
            }
        }
    }



}
