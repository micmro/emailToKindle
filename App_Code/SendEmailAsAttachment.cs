using System;
using System.Collections.Generic;
using System.Configuration;
using System.Text;
using System.Text.RegularExpressions;
using System.Net;
using System.Net.Mail;

using OpenPop.Pop3;

/// <summary>
/// Fetches all mail (based on configuration)
/// relies on OpenPop.net (http://hpop.sourceforge.net/) to use pop3 to retrieve emails
/// </summary>
public class SendEmailAsAttachment
{
    
    /// <summary>
    /// Fetches emails from server specified in the AppSettings, converts their body to a rtf or html attachment and sends it to the receiverEmailAddress
    /// </summary>
    /// <param name="receiverEmailAddress">Email address to send the email to</param>
    /// <returns>Emails retrieved from IncomingMailserverPop3Hostname</returns>
    public static List<OpenPop.Mime.Message> FetchMessagesAndSendAsAttachment(string receiverEmailAddress)
    {
        List<OpenPop.Mime.Message> messages = FetchAllMessages(ConfigurationManager.AppSettings["IncomingMailserverPop3Hostname"]
                        , int.Parse(ConfigurationManager.AppSettings["IncomingMailserverPop3Port"])
                        , false
                        , ConfigurationManager.AppSettings["IncomingMailserverUsername"]
                        , ConfigurationManager.AppSettings["IncomingMailserverPassword"]
                        , (ConfigurationManager.AppSettings["DeleteEmailsAfterFetching"].ToLower() == "true"));

        foreach (OpenPop.Mime.Message msg in messages)
        {

            MailMessage receivedMessage = msg.ToMailMessage();
            MailMessage messageToSend = new MailMessage(ConfigurationManager.AppSettings["EmailFrom"], receiverEmailAddress, "convert", "Send To Kindle. Date: " + DateTime.Now.ToShortDateString());

            string attachmentFileNameBase = Regex.Replace(receivedMessage.Subject, @"[^ a-zA-Z0-9_.-]", "_");
            if (receivedMessage.IsBodyHtml)
            {
                messageToSend.Attachments.Add(Attachment.CreateAttachmentFromString(receivedMessage.Body, attachmentFileNameBase + ".html", Encoding.UTF8, "text/html"));
            }
            else
            {
                messageToSend.Attachments.Add(Attachment.CreateAttachmentFromString(PlainTextToRtf(receivedMessage.Body), attachmentFileNameBase + ".rtf", Encoding.ASCII, "application/rtf"));
            }
            SmtpClient smtpclient = new SmtpClient(ConfigurationManager.AppSettings["OutgoingMailserverSmtpHostname"], int.Parse(ConfigurationManager.AppSettings["OutgoingMailserverSmtoPort"]));
            smtpclient.EnableSsl = false;
            smtpclient.Credentials = new NetworkCredential(ConfigurationManager.AppSettings["OutgoingMailserverUsername"], ConfigurationManager.AppSettings["OutgoingMailserverPassword"]);
            smtpclient.DeliveryMethod = SmtpDeliveryMethod.Network;

            smtpclient.Send(messageToSend);
        }

        return messages;
    }





    /// <summary>
    /// fetch all messages from a POP3 server
    /// </summary>
    /// <param name="hostname">Hostname of the server. For example: pop3.live.com</param>
    /// <param name="port">Host port to connect to. Normally: 110 for plain POP3, 995 for SSL POP3</param>
    /// <param name="useSsl">Whether or not to use SSL to connect to server</param>
    /// <param name="username">Username of the user on the server</param>
    /// <param name="password">Password of the user on the server</param>
    /// <param name="deleteImagAfterFetch">If true images on Pop3 will be deleted after beeing fetched</param>
    /// <returns>All Messages on the POP3 server</returns>
    public static List<OpenPop.Mime.Message> FetchAllMessages(string hostname, int port, bool useSsl, string username, string password, bool deleteImagAfterFetch)
    {
        // The client disconnects from the server when being disposed
        using (Pop3Client client = new Pop3Client())
        {
            // Connect to the server
            client.Connect(hostname, port, useSsl);
            client.Authenticate(username, password);
            int messageCount = client.GetMessageCount();

            // We want to download all messages
            List<OpenPop.Mime.Message> allMessages = new List<OpenPop.Mime.Message>(messageCount);
            for (int i = messageCount; i > 0; i--)
            {
                allMessages.Add(client.GetMessage(i));

                if (deleteImagAfterFetch)
                {
                    client.DeleteMessage(i);
                }
            }
            return allMessages;
        }
    }

    /// <summary>
    /// Converts plain (not html) email textto rtf
    /// </summary>
    /// <param name="plainText">plain email text</param>
    /// <returns>rtf escaped text (using Helvetica)</returns>
    private static string PlainTextToRtf(string plainText)
    {
        string rtf = @"{\rtf1\ansi{\fonttbl\f0\fswiss Helvetica;}\f0\pard ";
        rtf += GetRtfUnicodeEscapedString(plainText).Replace(Environment.NewLine, @" \par ");
        rtf += " }";
        return rtf;
    }


    //based on http://stackoverflow.com/questions/1368020/how-to-output-unicode-string-to-rtf-using-c
    /// <summary>
    /// converts charactes in string to unicode
    /// </summary>
    /// <param name="s">unicode string</param>
    /// <returns>rtf escaped string</returns>
    private static string GetRtfUnicodeEscapedString(string s)
    {
        var sb = new StringBuilder();
        foreach (var c in s)
        {
            if (c <= 0x7f)
            {
                sb.Append(c);
            }
            else
            {
                sb.Append("\\u" + Convert.ToUInt32(c) + "?");
            }
        }
        return sb.ToString();
    }
}