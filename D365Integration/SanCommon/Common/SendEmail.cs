using System;
using System.Collections.Generic;
using System.Net.Mail;
using System.Text;

// <copyright file="SendEmail.cs" company="JBS USA">
// Copyright (c) 2017 All Rights Reserved
// </copyright>
// <author>JBS USA</author>
// <date>01/11/2018 03:12:46 PM</date>
// <summary>Implements the SendEmail Plugin.</summary>

namespace SanCommon.Common
{
    public class SendEmail
    {
        private string _smtpHost;
        private int _smtpPort;
        private string _smtpDomain;
        private string _smtpUser;
        private string _smtpPassword;
        private string _smtpCcuser;

        public SendEmail(EmailConf emailConf)
        {
            this._smtpHost = emailConf.SmtpHost;
            this._smtpPort = emailConf.SmtpPort;
            this._smtpDomain = emailConf.SmtpDomain;
            this._smtpUser = emailConf.SmtpUser;
            this._smtpPassword = emailConf.SmtpPassword;
            this._smtpCcuser = emailConf.SmtpCcuser;
        }

        public void Send(MailAddress from, List<MailAddress> tos, List<MailAddress> ccs, string subject, string body, MailPriority priority, bool isHtml)
        {
            SmtpClient smtpClient = new SmtpClient();
            smtpClient.Host = _smtpHost;
            smtpClient.Port = _smtpPort;
            smtpClient.Credentials = new System.Net.NetworkCredential(_smtpUser, _smtpPassword, _smtpDomain);
            /*
            smtpClient.Host = "172.31.5.42";
            smtpClient.Port = 25;
            smtpClient.Credentials = new System.Net.NetworkCredential("administrator", "12125992575", "redacinc");
            */

            MailMessage message = new MailMessage();
            message.From = from ?? new MailAddress("crm@redacinc.com");

            if (tos.Count > 0)
            {
                foreach (MailAddress to in tos)
                {
                    message.To.Add(to);
                }
            }

            // Cc:
            message.CC.Add(_smtpCcuser);
            if (ccs != null)
            {
                foreach (MailAddress cc in ccs)
                {
                    message.CC.Add(cc);
                }
            }

            message.Subject = subject;

            message.IsBodyHtml = isHtml;
            message.Priority = priority;

            StringBuilder sb = new StringBuilder();
            if (isHtml) sb.Append("<html><body>");
            sb.Append(body);
            if (isHtml) sb.Append("</body></html>");

            message.Body = sb.ToString();

            smtpClient.Send(message);
        }

        public void SendException(Exception ex, string subjectAppend)
        {
            SmtpClient smtpClient = new SmtpClient();
            smtpClient.Host = _smtpHost;
            smtpClient.Port = _smtpPort;
            smtpClient.Credentials = new System.Net.NetworkCredential(_smtpUser, _smtpPassword, _smtpDomain);
            /*
            smtpClient.Host = "172.31.5.42";
            smtpClient.Port = 25;
            smtpClient.Credentials = new System.Net.NetworkCredential("administrator", "12125992575", "redacinc");
            */

            MailMessage message = new MailMessage();
            message.From = new MailAddress("crm@redacinc.com");

            //message.To.Add(new MailAddress("ae@multinet-usa.com"));
            message.To.Add(new MailAddress("support-redac@jbs.com"));
            message.CC.Add(new MailAddress(_smtpCcuser));

            message.Subject = "RedacImportCsv Exception Occurred! - " + subjectAppend;

            message.IsBodyHtml = true;

            StringBuilder sb = new StringBuilder();
            sb.Append("<html><body>");
            sb.Append("<p>RedacImportCsvにおいて、以下の例外が発生しました。</p>");
            sb.Append("<hr />");
            sb.Append("<p>");
            sb.Append(ex.ToString());
            sb.Append("</p>");
            sb.Append("<hr />");
            sb.Append("</body></html>");

            message.Body = sb.ToString();

            smtpClient.Send(message);
        }
    }

    public class EmailConf
    {
        private static string _smtpHost;
        private static int _smtpPort;
        private static string _smtpDomain;
        private static string _smtpUser;
        private static string _smtpPassword;
        private static string _smtpCcuser;

        public EmailConf()
        {
            SmtpHost = ImportCommon.GetAppSetting("SmtpHost");
            SmtpPort = Int32.Parse(ImportCommon.GetAppSetting("SmtpPort"));
            SmtpDomain = ImportCommon.GetAppSetting("SmtpDomain");
            SmtpUser = ImportCommon.GetAppSetting("SmtpUser");
            SmtpPassword = ImportCommon.GetAppSetting("SmtpPassword");
            SmtpCcuser = ImportCommon.GetAppSetting("SmtpCcuser");
        }

        public string SmtpHost
        {
            get
            {
                return _smtpHost;
            }
            set
            {
                _smtpHost = value;
            }
        }

        public int SmtpPort
        {
            get
            {
                return _smtpPort;
            }
            set
            {
                _smtpPort = value;
            }
        }

        public string SmtpDomain
        {
            get
            {
                return _smtpDomain;
            }
            set
            {
                _smtpDomain = value;
            }
        }

        public string SmtpUser
        {
            get
            {
                return _smtpUser;
            }
            set
            {
                _smtpUser = value;
            }
        }

        public string SmtpPassword
        {
            get
            {
                return _smtpPassword;
            }
            set
            {
                _smtpPassword = value;
            }
        }

        public string SmtpCcuser
        {
            get
            {
                return _smtpCcuser;
            }
            set
            {
                _smtpCcuser = value;
            }
        }
    }
}
                                
                                
                                
                                
                                
                                
                                
                                
