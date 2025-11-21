//20250326001-保險公司受理前十大商品年報排程 20250507 by Harrison
using Top10ReportProcess.Model;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Net.Mail;
using System.Net;
using NLog;

namespace Top10ReportProcess
{
    public class MailHelper
    {
        public string ExecEnv = ConfigurationManager.AppSettings["ENV"];
        public string LogMailAdd = ConfigurationManager.AppSettings["LogMailAdd"];
        public string SysMailAdd = ConfigurationManager.AppSettings["SysMailAdd"];
        private static readonly NLog.Logger logger = NLog.LogManager.GetCurrentClassLogger();

        public AutoMailInfo GetAutoMailInfo()
        {
            AutoMailInfo rtnMailInfo = new AutoMailInfo();
            rtnMailInfo.SenderAddress = ConfigurationManager.AppSettings["SenderAddress"];
            rtnMailInfo.SenderDisplayname = ConfigurationManager.AppSettings["SenderDisplayname"];
            rtnMailInfo.SmtpAddress = ConfigurationManager.AppSettings["SmtpAddress"];
            rtnMailInfo.SmtpPort = int.Parse(ConfigurationManager.AppSettings["SmtpPort"]);
            rtnMailInfo.SmtpCredentials = ConfigurationManager.AppSettings["SmtpCredentials"];
            rtnMailInfo.CredentialsId = ConfigurationManager.AppSettings["CredentialsId"];
            rtnMailInfo.CredentialsPwd = ConfigurationManager.AppSettings["CredentialsPwd"];

            return rtnMailInfo;
        }

        /// <summary>
        /// Mail夾附件發送
        /// </summary>
        /// <param name="reportStream"></param>
        /// <param name="fileName"></param>
        /// <param name="Subject"></param>
        /// <param name="MailADD"></param>
        public void SendReportByEmail(MemoryStream reportStream, string fileName, string Subject, List<AutoEmailMailaddress> MailADD)
        {
            SmtpClient client = new SmtpClient();
            AutoMailInfo Logmailinfo = GetAutoMailInfo();
            using (MailMessage lo_mm = new MailMessage())
            {
                lo_mm.From = new MailAddress(Logmailinfo.SenderAddress, Logmailinfo.SenderDisplayname);
                lo_mm.Subject = Subject;
                if (ExecEnv == "T")
                {
                    lo_mm.Subject = "[TEST]" + lo_mm.Subject;
                    lo_mm.To.Add(new MailAddress("harrison831123@mail.everprobks.com.tw"));
                }
                else
                {
                    foreach(var item in MailADD)
                    {
                        lo_mm.To.Add(new MailAddress(item.mail_address));
                    }
                }

                lo_mm.SubjectEncoding = System.Text.Encoding.UTF8;
                lo_mm.BodyEncoding = System.Text.Encoding.UTF8;
                lo_mm.IsBodyHtml = true;
                lo_mm.Priority = System.Net.Mail.MailPriority.High;

                // 加入 MemoryStream 作為附件
                var attachment = new Attachment(reportStream, fileName, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
                lo_mm.Attachments.Add(attachment);

                client.Host = Logmailinfo.SmtpAddress; //設定smtp Server
                client.Port = Logmailinfo.SmtpPort; //設定Port
                client.UseDefaultCredentials = true;
                client.EnableSsl = true; //gmail預設開啟驗證

                // Check Sender Mail ID and PWD
                if (Logmailinfo.SmtpCredentials == "Y")
                {
                    client.Credentials = new NetworkCredential(Logmailinfo.CredentialsId, Logmailinfo.CredentialsPwd);
                }
                client.Send(lo_mm);
            }
            client.Dispose();
        }

        public void SendLogMail(String FundID = "")
        {
            var logfile = (LogManager.Configuration.FindTargetByName("File") as NLog.Targets.FileTarget).FileName
                .Render(new LogEventInfo() { TimeStamp = DateTime.Now, LoggerName = "loggerName" });

            (LogManager.Configuration.FindTargetByName("File") as NLog.Targets.FileTarget).KeepFileOpen = false;
            SmtpClient client = new SmtpClient();
            try
            {
                AutoMailInfo Logmailinfo = GetAutoMailInfo();
                using (MailMessage lo_mm = new MailMessage())
                {
                    lo_mm.From = new MailAddress(Logmailinfo.SenderAddress, Logmailinfo.SenderDisplayname);
                    lo_mm.Subject = String.Format("Top10ReportProcess_異常通知{0}", FundID);
                    if (ExecEnv == "T")
                    {
                        lo_mm.Subject = "[TEST]" + lo_mm.Subject;
                        lo_mm.To.Add(new MailAddress("harrison831123@mail.everprobks.com.tw"));
                    }
                    else
                    {
                        lo_mm.To.Add(new MailAddress(LogMailAdd));
                    }

                    lo_mm.SubjectEncoding = System.Text.Encoding.UTF8;
                    lo_mm.BodyEncoding = System.Text.Encoding.UTF8;
                    lo_mm.IsBodyHtml = true;
                    lo_mm.Priority = System.Net.Mail.MailPriority.High;
                    lo_mm.Attachments.Add(new Attachment(logfile));

                    client.Host = Logmailinfo.SmtpAddress; //設定smtp Server
                    client.Port = Logmailinfo.SmtpPort; //設定Port
                    client.UseDefaultCredentials = true;
                    client.EnableSsl = true; //gmail預設開啟驗證

                    // Check Sender Mail ID and PWD
                    if (Logmailinfo.SmtpCredentials == "Y")
                    {
                        client.Credentials = new NetworkCredential(Logmailinfo.CredentialsId, Logmailinfo.CredentialsPwd);
                    }
                    client.Send(lo_mm);
                }
            }
            catch (Exception ex)
            {
                logger.Error(ex.Message);
            }
            finally
            {
                client.Dispose();
            }
        }

    }

}
