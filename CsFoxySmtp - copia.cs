using System;
using System.Runtime.InteropServices;
using System.Net.Mail;
using System.Net.Mime;
using System.Collections.Generic;
using System.IO;

namespace CsFoxySmtp
{
    [ClassInterface(ClassInterfaceType.AutoDual)]
    [ProgId("CsFoxySmtp")]

    /// <summary>
    /// DLL en C# para Enviar Correo desde VFP.
    /// Microsoft: Outlook y Office365
    /// Google: Gmail
    /// </summary>

    public class CsFoxySmtp : System.EnterpriseServices.ServicedComponent
    {
        private readonly string _version = "0.8.6";
        public string from     = "";
        public string user     = "";
        public string password = "";
        public string subjet   = "";
        public string body     = "";
        public bool   bodyHtml = true;
        public bool   clear    = false;
        public bool   fileSize = false;
        public bool   notification = false;
        private Attachment fileAdd;

        private readonly List<string> _emailTo   = new List<string>();
        private readonly List<string> _emailCc   = new List<string>();
        private readonly List<string> _emailBcc  = new List<string>();
        private readonly List<string> _filesList = new List<string>();

        public string server   = "";
        public Int32  port     = 587;
        public bool   ssl      = true;
        public int    priority = 0;
        private MailMessage mailMessage = new MailMessage();

        public string Version => _version;

        public string Error { get; private set; } = "";

        public decimal Size { get; private set; } = 0;

        public void Clear()
        {
            from = "";
            subjet = "";
            body = "";
            Error = "";
            Size = 0;
            ClearTo();
            ClearCc();
            ClearBcc();
            ClearAttachments();
        }

        private MailPriority Priority(Int32 num)
        {
            switch (num)
            {
                case 0:
                    return MailPriority.Normal;
                case 1:
                    return MailPriority.Low;
                case 2:
                    return MailPriority.High;
                default:
                    return MailPriority.Normal;
            }
        }

        public Boolean Smtp()
        {
            bool send = false;
            Error = "";

            mailMessage = new MailMessage(
                          from: from,
                          to: To,
                          subject: subjet,
                          body: body
                          )
            {
                IsBodyHtml = bodyHtml,
                BodyEncoding = System.Text.Encoding.UTF8,
                SubjectEncoding = System.Text.Encoding.UTF8,
                Priority = Priority(priority),
            };

            if ( notification == true )
            {
                mailMessage.Headers.Add( "Disposition-Notification-To", user );
                mailMessage.Headers.Add( "Return-Receipt-To", user );
            }

            if (_emailCc.Count > 0)
            {
                _emailCc.ForEach(delegate (string email)
                {
                    mailMessage.CC.Add(email);
                });
            };

            if (_emailBcc.Count > 0)
            {
                _emailBcc.ForEach(delegate (string email)
                {
                    mailMessage.Bcc.Add(email);
                });
            };

            if (_filesList.Count > 0)
            {
                _filesList.ForEach(delegate (string file)
                {
                    fileAdd = new Attachment(file, MediaTypeNames.Application.Octet);
                    mailMessage.Attachments.Add(fileAdd);
                });
            };

            SmtpClient smtpClient = new SmtpClient(server)
            {
                EnableSsl = ssl,
                Host = server,
                Port = port,
                Credentials = new System.Net.NetworkCredential(user, password),
                //TargetName = "STARTTLS/smtp.office365.com",
                //UseDefaultCredentials = false,
                //DeliveryMethod = SmtpDeliveryMethod.Network,
            };

            try
            {
                smtpClient.Send(mailMessage);

                if ( clear == true )
                {
                    Clear();
                }
            }
            catch (Exception ex)
            {
                Error = ex.ToString();
            }

            smtpClient.Dispose();

            if (_filesList.Count > 0)
            {
                fileAdd.Dispose();
            }

            mailMessage.Attachments.Clear();

            if ( Error.Length <= 0 )
            {
                send = true;
            }
            
            return send;
        }

        public string Attachments => ListToString(_filesList);
        public decimal AddAttachments(string pathFile)
        {
            _filesList.Add(pathFile);
            if ( fileSize == true )
            {
                decimal size = FileSize(pathFile);
                Size += size;
            } else
            {
                Size = 0;
            }
            return Size;
        }
        public void ClearAttachments() => _filesList.Clear();
        public int CountAttachments => _filesList.Count;

        public string Cc => ListToString(_emailCc);
        public void AddCc(string email) => _emailCc.Add(email);
        public void ClearCc() => _emailCc.Clear();
        public int CountCc => _emailCc.Count;

        public string Bcc => ListToString(_emailBcc);
        public void AddBcc(string email) => _emailBcc.Add(email);
        public void ClearBcc() => _emailBcc.Clear();
        public int CountBcc => _emailBcc.Count;

        public string To => ListToString(_emailTo);
        public void AddTo(string email) => _emailTo.Add(email);
        public void ClearTo() => _emailTo.Clear();
        public int CountTo => _emailTo.Count;


        private string ListToString(List<string> list)
        {
            string listToString = "";
            int index = 0;
            list.ForEach(delegate (string item)
            {
                index += 1;
                if (index > 1)
                {
                    listToString += ", ";
                }
                listToString += item;
            });
            return listToString;
        }

        private decimal FileSize(string pathFile)
        {
            FileInfo file = new FileInfo(pathFile);
            return (file.Length / 1024);
        }
    }

}
