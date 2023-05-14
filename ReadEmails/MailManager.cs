using OpenPop.Mime.Decode;
using OpenPop.Mime;
using OpenPop.Pop3;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ReadEmails
{
    public class MailManager
    {
        public string Host { get; set; }
        public string Email { get; set; }
        public string PasswordFile { get; set; }
        public int Port { get; set; }
        public bool Ssl { get; set; }
        public List<Message> MessageList { get; set; }

        public MailManager(string host, string email, string passwordFile, int port, bool ssl)
        {
            Host = host;
            Email = email;
            PasswordFile = passwordFile;
            Port = port;
            Ssl = ssl;
            MessageList = GetEmails();
        }

        public List<Message> GetEmails()
        {
            string password = File.ReadAllText(PasswordFile);
            try
            {
                using (Pop3Client _client = new())
                {
                    _client.Connect(Host, Port, Ssl);
                    _client.Authenticate(Email, password);

                    int messageCount = _client.GetMessageCount();

                    var messages = new List<Message>(messageCount);

                    int mailsToRead;
                    if (messageCount > 50) mailsToRead = messageCount - 50;
                    else mailsToRead = messageCount;

                    for (int i = mailsToRead; i < messageCount; i++)
                    {
                        Message message = _client.GetMessage(i + 1);
                        messages.Add(message);
                    }

                    return messages;
                }
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
        }

        public List<Message> FilterBySender(string sender)
        {
            return MessageList.Where(x => x.Headers.From.Address == sender).ToList();
        }

        public List<Message> FilterBySubject(string subject)
        {
            return MessageList.Where(x => x.Headers.Subject.ToLower()
                .Contains(subject.ToLower()))
                .ToList();
        }

        public List<Message> FilterBySenderAndSubject(string sender, string subject)
        {
            return MessageList.Where(x => x.Headers.From.Address == sender &&
                x.Headers.Subject.ToLower()
                .Contains(subject.ToLower()))
                .ToList();
        }

        public void DownloadAttachment(List<Message> messageList, string destinationFolder, string fileNameWithExtension)
        {
            if (!messageList.Any()) throw new Exception("Empty list of emails");

            var lastMessage = messageList.Last();
            var attachments = lastMessage.FindAllAttachments();

            if (!attachments.Any()) throw new Exception("No attachments found");

            foreach (var attachment in attachments)
            {
                string filePath = Path.Combine(destinationFolder, attachment.FileName);

                if (File.Exists(filePath)) throw new Exception("File already exists");

                if (attachment.FileName.Equals(fileNameWithExtension))
                {
                    FileStream Stream = new FileStream(filePath, FileMode.Create);
                    BinaryWriter BinaryStream = new BinaryWriter(Stream);
                    BinaryStream.Write(attachment.Body);
                    BinaryStream.Close();
                }
            }
        }
    }
}
