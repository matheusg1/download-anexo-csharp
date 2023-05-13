using OpenPop.Mime;
using OpenPop.Mime.Decode;
using OpenPop.Pop3;
using ReadEmails;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using Xunit;

namespace ReadEmailTests
{
    public class MailTests
    {
        [Fact]
        public void ConfigureEncoding()
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
            EncodingFinder.AddMapping("ks_c_5601-1987", Encoding.GetEncoding(949));
        }

        [Theory]
        [InlineData("outlook.office365.com", "tecmatheusg@hotmail.com", 995, true, @"C:\Users\Matheus\Desktop\Senha.txt", "matheusmt021@gmail.com", 3)]
        [InlineData("outlook.office365.com", "tecmatheusg@hotmail.com", 995, true, @"C:\Users\Matheus\Desktop\Senha.txt", "marketing@ionic.io", 2)]
        [InlineData("outlook.office365.com", "tecmatheusg@hotmail.com", 995, true, @"C:\Users\Matheus\Desktop\Senha.txt", "mts_gomes_silva@hotmail.com", 0)]
        public void NumberOfMailsTest(string host, string email, int port, bool ssl, string passwordfile, string sender, int numberEmails)
        {
            //arrange
            var mM = new MailManager(host, email, passwordfile, port, ssl);

            //act
            var filteredMails = mM.FilterBySender(sender);

            //assert
            Assert.Equal(numberEmails, filteredMails.Count);
        }

        [Theory]
        [InlineData("outlook.office365.com", "tecmatheusg@hotmail.com", 995, true, @"C:\Users\Matheus\Desktop\Senha.txt", "img3")]
        [InlineData("outlook.office365.com", "tecmatheusg@hotmail.com", 995, true, @"C:\Users\Matheus\Desktop\Senha.txt", "img2")]
        public void FindEmailBySubjectTest(string host, string email, int port, bool ssl, string passwordfile, string subject)
        {
            //arrange
            var mM = new MailManager(host, email, passwordfile, port, ssl);

            //act
            var filteredMails = mM.FilterBySubject(subject);

            //assert
            Assert.Single(filteredMails);
        }

        [Theory]
        [InlineData("outlook.office365.com", "tecmatheusg@hotmail.com", 995, true, @"C:\Users\Matheus\Desktop\Senha.txt", "matheusmt021@gmail.com","img3", 1)]
        [InlineData("outlook.office365.com", "tecmatheusg@hotmail.com", 995, true, @"C:\Users\Matheus\Desktop\Senha.txt", "anything@gmail.com", "img2", 0)]
        public void FindEmailBySenderAndSubjectTest(string host, string email, int port, bool ssl, string passwordfile, string sender, string subject, int result)
        {
            //arrange
            var mM = new MailManager(host, email, passwordfile, port, ssl);

            //act
            var filteredMails = mM.FilterBySenderAndSubject(sender, subject);

            //assert
            Assert.Equal(result, filteredMails.Count);
        }
        [Theory]
        [InlineData("outlook.office365.com", "tecmatheusg@hotmail.com", 995, true, @"C:\Users\Matheus\Desktop\Senha.txt", "matheusmt021@gmail.com", "img2")]        
        public void DownloadAttachmentTest(string host, string email, int port, bool ssl, string passwordfile, string sender, string subject)
        {
            //arrange
            var mM = new MailManager(host, email, passwordfile, port, ssl);

            //act
            var filteredMails = mM.FilterBySenderAndSubject(sender, subject);

            string destinationFolder = @"C:\Users\Matheus\Desktop";
            string fileNameWithExtension = "1222.png";

            //assert
            mM.DownloadAttachment(filteredMails, destinationFolder, fileNameWithExtension);
            Assert.Equal(1, 1);            
        }

        [Theory]
        [InlineData("outlook.office365.com", "tecmatheusg@hotmail.com", 995, true, @"C:\Users\Matheus\Desktop\Senha.txt", "matheusmt021@gmail.com", "img3")]
        [InlineData("outlook.office365.com", "tecmatheusg@hotmail.com", 995, true, @"C:\Users\Matheus\Desktop\Senha.txt", "anything@gmail.com", "img2")]
        public void DownloadAttachmentExceptionTest(string host, string email, int port, bool ssl, string passwordfile, string sender, string subject)
        {
            //arrange
            var mM = new MailManager(host, email, passwordfile, port, ssl);

            //act
            var filteredMails = mM.FilterBySenderAndSubject(sender, subject);

            string destinationFolder = @"C:\Users\Matheus\Desktop";
            string fileNameWithExtension = "1222.png";

            //assert          
            Assert.Throws<Exception>(() => mM.DownloadAttachment(filteredMails, destinationFolder, fileNameWithExtension));
        }

    }
}