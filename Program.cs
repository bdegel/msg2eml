using System;
using System.IO;
using MsgReader.Outlook;
using MimeKit;

namespace msg2eml
{
    internal class Program
    {
        static void Main(string[] args)
        {
            if (args.Length > 0)
            {
                ConvertMsgToEml(args[0], args.Length > 1 ? args[1] : Path.ChangeExtension(args[0], ".eml"));
            }
            else
            {
                Console.WriteLine("Usage: msg2eml <input.msg> [output.eml]");
                Console.WriteLine("If output.eml is not specified, it will be created in the same directory as input.msg with the same name.");
            }
        }
        static void ConvertMsgToEml(string msgPath, string emlPath)
        {
            try
            {
                using (var msg = new Storage.Message(msgPath))
                {
                    var mime = new MimeMessage();

                    // From
                    if (!string.IsNullOrWhiteSpace(msg.Sender.Email))
                        mime.From.Add(MailboxAddress.Parse(msg.Sender.Email));

                    // To/Cc
                    foreach (var to in msg.Recipients)
                        mime.To.Add(MailboxAddress.Parse(to.Email));

                    foreach (var cc in msg.Recipients)
                        mime.Cc.Add(MailboxAddress.Parse(cc.Email));

                    mime.Subject = msg.Subject ?? "";

                    // Body: check if HTML exists, else text
                    var bodyBuilder = new BodyBuilder();
                    if (!string.IsNullOrWhiteSpace(msg.BodyHtml))
                        bodyBuilder.HtmlBody = msg.BodyHtml;
                    else
                        bodyBuilder.TextBody = msg.BodyText ?? "";

                    // Attachments
                    if (msg.Attachments != null)
                    {
                        foreach (Storage.Message.Attachment att in msg.Attachments)
                        {
                            bodyBuilder.Attachments.Add(att.FileName, att.Data);
                        }
                    }

                    mime.Body = bodyBuilder.ToMessageBody();

                    // Write EML
                    using (var fs = File.Create(emlPath))
                    {
                        mime.WriteTo(fs);
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error converting '{msgPath}' to '{emlPath}': {ex.Message}");
            }
        }
    }
}
