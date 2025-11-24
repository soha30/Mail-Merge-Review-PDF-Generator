using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Net;
using System.Net.Mail;
using System.Threading;
using System.IO;
using System.Configuration;
using System.Web.UI.WebControls;

namespace ali1982ReviewMailmerge.App_Code
{
/// <summary>
/// MailManager class for sending emails using Gmail SMTP.
/// </summary>
public class MailManager
    {
        // Public fields for email configuration
        public string FromEmail;
        public string ToEmail;
        public string Subject;
        public string Body;
        public bool IsBodyHtml;
        public string SmtpHost;
        public int SmtpPort;
        public bool EnableSsl;
        public string SmtpUserName;
        public string SmtpPassword;

        // Constructor to initialize fields from web.config
        public MailManager()
        {
            // Initialize fields from web.config settings
            FromEmail = ConfigurationManager.AppSettings["emailFrom"];
            ToEmail = ConfigurationManager.AppSettings["emailTo"];
            SmtpHost = ConfigurationManager.AppSettings["HostsmtpAddress"];
            SmtpPort = int.Parse(ConfigurationManager.AppSettings["PortNumber"]);
            EnableSsl = bool.Parse(ConfigurationManager.AppSettings["EnableSSL"]);
            SmtpUserName = ConfigurationManager.AppSettings["emailUserName"];
            SmtpPassword = ConfigurationManager.AppSettings["emailPassword"];

            // Default values for subject and body
            Subject = "Notification of Site Activity via Gmail SMTP";
            IsBodyHtml = true;
            Body = @"This is the default body. Replace it with your actual message.";
        }

        // Method to send a simple email
        public string SendEmail()
        {
            try
            {
                // Create a MailMessage object with the specified fields
                using (MailMessage mail = new MailMessage(FromEmail, ToEmail, Subject, Body))
                {
                    mail.IsBodyHtml = IsBodyHtml;

                    // Configure the SMTP client
                    SmtpClient smtpClient = new SmtpClient(SmtpHost, SmtpPort)
                    {
                        Credentials = new NetworkCredential(SmtpUserName, SmtpPassword),
                        EnableSsl = EnableSsl
                    };

                    // Send the email
                    smtpClient.Send(mail);
                    return "Email sent successfully";
                }
            }
            catch (Exception ex)
            {
                // Return the error message if an exception occurs
                return "An error occurred: " + ex.Message;
            }
        }

        // Method to send an email with an attachment
        public string SendEmailWithAttachment(FileUpload fileUpload)
        {
            try
            {
                // Create a MailMessage object with the specified fields
                using (MailMessage mail = new MailMessage(FromEmail, ToEmail, Subject, Body))
                {
                    mail.IsBodyHtml = IsBodyHtml;

                    // Attach files if any are uploaded
                    if (fileUpload.HasFile)
                    {
                        foreach (HttpPostedFile file in fileUpload.PostedFiles)
                        {
                            string fileName = Path.GetFileName(file.FileName);
                            mail.Attachments.Add(new Attachment(file.InputStream, fileName));
                        }
                    }

                    // Configure the SMTP client
                    SmtpClient smtpClient = new SmtpClient(SmtpHost, SmtpPort)
                    {
                        Credentials = new NetworkCredential(SmtpUserName, SmtpPassword),
                        EnableSsl = EnableSsl
                    };

                    // Send the email with attachments
                    smtpClient.Send(mail);
                    return "Email with attachment sent successfully";
                }
            }
            catch (Exception ex)
            {
                // Return the error message if an exception occurs
                return "An error occurred: " + ex.Message;
            }
        }

        // Method to send a custom email to a specified recipient with a custom subject and body
        public string SendCustomEmail(string toEmail, string subject, string body)
        {
            try
            {
                // Create a MailMessage object with custom fields
                using (MailMessage mail = new MailMessage(FromEmail, toEmail, subject, body))
                {
                    mail.IsBodyHtml = IsBodyHtml;

                    // Configure the SMTP client
                    SmtpClient smtpClient = new SmtpClient(SmtpHost, SmtpPort)
                    {
                        Credentials = new NetworkCredential(SmtpUserName, SmtpPassword),
                        EnableSsl = EnableSsl
                    };

                    // Send the custom email
                    smtpClient.Send(mail);
                    return "Custom email sent successfully";
                }
            }
            catch (Exception ex)
            {
                // Return the error message if an exception occurs
                return "An error occurred: " + ex.Message;
            }
        }
    }

}
