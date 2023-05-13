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
            SetupEncoding();

            var manager = new MailManager();
            manager._passwordFile = @"C:\Users\Matheus\Desktop\Senha.txt";
            manager._host = "outlook.office365.com";
            manager._email = "tecmatheusg@hotmail.com";
            manager._port = 995;
            manager._ssl = true;

            string sender = "matheusmt021 @gmail.com";
            //string subject = "img2";

            List<Message> messageList = manager.GetEmails();

            List<Message> filteredMessages = manager.FilterBySender(sender);                        
           
            string destinationFolder = @"C:\Users\Matheus\Desktop";
            string fileNameWithExtension = "1222.png";

            manager.DownloadAttachment(filteredMessages, destinationFolder, fileNameWithExtension);
        }
        public static void SetupEncoding()
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
            EncodingFinder.AddMapping("ks_c_5601-1987", Encoding.GetEncoding(949));
        }
    }
}
