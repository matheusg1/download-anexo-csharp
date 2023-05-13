using OpenPop.Mime.Decode;
using OpenPop.Mime;
using OpenPop.Pop3;
using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.Intrinsics.X86;
using System.Text;
using System.Linq;

namespace ReadEmails
{
    public class Program
    {
        static void Main(string[] args)
        {            
            string passwordFile = @"C:\Users\Matheus\Desktop\Senha.txt";
            string host = "outlook.office365.com";
            string email = "tecmatheusg@hotmail.com";
            string sender = "matheusmt021@gmail.com";
            var port = 995;
            var ssl = true;

            List<Message> messageList = GetEmailsFromSender(host, port, email, ssl, passwordFile, sender);

            string destinationFolder = @"C:\Users\Matheus\Desktop";
            string fileNameWithExtension = "1222.png";

            DownloadAttachment(messageList, destinationFolder, fileNameWithExtension);
        }

        public static List<Message> GetEmailsFromSender(string host, int port, string email, bool ssl, string passwordFile, string sender)
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
            EncodingFinder.AddMapping("ks_c_5601-1987", Encoding.GetEncoding(949));

            string textFile = passwordFile;
            string password = File.ReadAllText(textFile);

            using (Pop3Client _client = new())
            {
                _client.Connect(host, port, ssl);
                _client.Authenticate(email, password);

                int messageCount = _client.GetMessageCount();

                var Messages = new List<Message>(messageCount);

                for (int i = messageCount - 50; i < messageCount; i++)
                {
                    Message message = _client.GetMessage(i + 1);
                    Messages.Add(message);
                }

                return Messages.Where(x => x.Headers.From.Address == sender).ToList();
            }
        }
        public static void DownloadAttachment(List<Message> messages, string destinationFolder, string fileNameWithExtension)
        {            
            var lastMessage = messages.Last();

            foreach (var attachment in lastMessage.FindAllAttachments())
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
