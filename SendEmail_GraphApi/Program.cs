using Microsoft.Graph;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SendEmail_GraphApi
{
    public class Program
    {
        public static void Main(string[] args)
        {
            Program pp = new Program();
            pp.SendEmail();
        }

        public async Task SendEmail()
        {
            // Arrange.
            GraphService graphService = new GraphService();
            string subject = "Test email from ASP.NET 4.6 Connect sample";
            string bodyContent = "<html><body>The body of the test email.</body></html>";
            List<Recipient> recipientList = new List<Recipient>();
            recipientList.Add(new Recipient
            {
                EmailAddress = new EmailAddress
                {
                    Address = "srinivas.ch@ensurity.com"
                }
            });
            Message message = new Message
            {
                Body = new ItemBody
                {
                    Content = bodyContent,
                    ContentType = BodyType.Html,
                },
                Subject = subject,
                ToRecipients = recipientList
            };

            // Act
            Task task = graphService.SendEmail(client, message);

            // Assert
            Task.WaitAll(task);
        }
    }
}
