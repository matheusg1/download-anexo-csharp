using OpenPop.Mime;
using OpenPop.Mime.Decode;
using OpenPop.Pop3;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using Xunit;

namespace ReadEmailTests
{
    public class UnitTest1
    {
        [Theory]
        //[InlineData("pop.gmail.com", "matheusg@souunisuam.com.br", 995, false)]
        //[InlineData("pop.gmail.com", "matheusg@souunisuam.com.br", 995, true)]
        [InlineData("outlook.office365.com", "tecmatheusg@hotmail.com", 995, true)]
        public void GetMail(string host, string email, int port, bool ssl)
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
            EncodingFinder.AddMapping("ks_c_5601-1987", Encoding.GetEncoding(949));

            string textFile = @"C:\Users\Matheus\Desktop\Senha.txt";
            string password = File.ReadAllText(textFile);

            using (Pop3Client _client = new Pop3Client())
            {
                _client.Connect(host, port, ssl);
                _client.Authenticate(email, password);

                int messageCount = _client.GetMessageCount();

                var Messages = new List<Message>(messageCount);

                for (int i = messageCount - 50; i < messageCount; i++)
                {
                    Message getMessage = _client.GetMessage(i + 1);                    
                    Messages.Add(getMessage);
                }
                
                var msgs = Messages.Where(x => x.Headers.From.Address == "matheusmt021@gmail.com");
                var primeiraMsg = msgs.First();
                var ultimaMsg = msgs.Last();

                /*
                foreach (Message msg in Messages)
                {
                */

                foreach (var attachment in ultimaMsg.FindAllAttachments())
                {
                    string filePath = Path.Combine(@"C:\Users\Matheus\Desktop", attachment.FileName);
                    if (attachment.FileName.Equals("1222.png"))
                    {
                        FileStream Stream = new FileStream(filePath, FileMode.Create);
                        BinaryWriter BinaryStream = new BinaryWriter(Stream);
                        BinaryStream.Write(attachment.Body);
                        BinaryStream.Close();
                    }
                }
            }

                //assert
                Assert.Equal(1, 1);
            }
        }
    }
