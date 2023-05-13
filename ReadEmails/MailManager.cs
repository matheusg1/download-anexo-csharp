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
        public string _host { get; set; }
        public string _email { get; set; }
        public string _passwordFile { get; set; }
        public int _port { get; set; }
        public bool _ssl { get; set; }
        public List<Message> _messageList { get; set; }

        public MailManager()
        {
        }

        public MailManager(string host, string email, string passwordFile, int port, bool ssl)
        {
            _host = host;
            _email = email;
            _passwordFile = passwordFile;
            _port = port;
            _ssl = ssl;
            _messageList = GetEmails();
        }

        public List<Message> GetEmails()
        {
            string password = File.ReadAllText(_passwordFile);
            try
            {
                using (Pop3Client _client = new())
                {
                    _client.Connect(_host, _port, _ssl);
                    _client.Authenticate(_email, password);

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
            return _messageList.Where(x => x.Headers.From.Address == sender).ToList();
        }

        public List<Message> FilterBySubject(string subject)
        {
            return _messageList.Where(x => x.Headers.Subject.ToLower()
                .Contains(subject.ToLower()))
                .ToList();
        }

        public List<Message> FilterBySenderAndSubject(string sender, string subject)
        {
            return _messageList.Where(x => x.Headers.From.Address == sender &&
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
