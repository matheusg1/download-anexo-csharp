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

            var passwordFile = @"C:\Users\Matheus\Desktop\Senha.txt";
            var host = "outlook.office365.com";
            var email = "tecmatheusg@hotmail.com";
            var port = 995;
            var ssl = true;
            var manager = new MailManager(host, email, passwordFile, port, ssl);

            string sender = "matheusmt021@gmail.com";
            //string subject = "img2";
            List<Message> filteredMessages = manager.FilterBySenderAndSubject(sender, "img2");                        
           
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
