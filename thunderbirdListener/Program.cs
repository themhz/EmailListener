using System;
using System.IO;
using System.Threading.Tasks;
using MailKit.Net.Imap;
using MailKit;
using MimeKit;
using Microsoft.Extensions.Configuration;
using System.Drawing.Printing;
using MimeKit;
using System.Drawing;

class Program
{
    static async Task Main(string[] args)
    {
        var config = LoadConfiguration();
        await ReadEmailsAsync(config);
    }

    private static IConfiguration LoadConfiguration()
    {
        var outputPath = AppDomain.CurrentDomain.BaseDirectory;
        var configPath = Path.Combine(outputPath, "config.ini");

        var builder = new ConfigurationBuilder()
            .SetBasePath(outputPath)
            .AddIniFile(configPath, optional: false, reloadOnChange: true);

        return builder.Build();
    }



    private static async Task ReadEmailsAsync(IConfiguration config)
    {
        // Read email credentials and IMAP settings from the configuration
        string email = config["Email"];
        string password = config["Password"];
        string imapServer = config["ImapServer"];
        int imapPort = int.Parse(config["ImapPort"]);


        using (var client = new ImapClient())
        {
            // Accept all SSL certificates (not recommended for production)
            client.ServerCertificateValidationCallback = (s, c, h, e) => true;

            // Connect to the IMAP server
            await client.ConnectAsync(imapServer, imapPort, true);

            // Authenticate with the server
            await client.AuthenticateAsync(email, password);

            // Open the Inbox folder
            var inbox = client.Inbox;
            await inbox.OpenAsync(FolderAccess.ReadOnly);

            int lastProcessedMessage = inbox.Count;

            // Monitor the mailbox for new messages
            while (true)
            {
                await client.NoOpAsync();

                int currentMessageCount = inbox.Count;

                if (currentMessageCount > lastProcessedMessage)
                {
                    for (int i = lastProcessedMessage; i < currentMessageCount; i++)
                    {
                        var message = await inbox.GetMessageAsync(i);                                               

                        if(message.From.Mailboxes.FirstOrDefault().Address == "notifications@ecwid.com" && message.Subject.Contains("Comanda"))
                        {
                            Console.WriteLine($"Message received:");
                            Console.WriteLine($"Subject: {message.Subject}");
                            Console.WriteLine($"From: {message.From}");
                            Console.WriteLine($"Date: {message.Date}");
                            Console.WriteLine($"Body: {message.TextBody}");
                            Console.WriteLine("-------------------------");

                            PrintMessageBody(message);
                        }
                    }

                    lastProcessedMessage = currentMessageCount;
                }

                await Task.Delay(TimeSpan.FromSeconds(30));
            }
        }
    }


    public static void PrintMessageBody(MimeMessage message)
    {
        if (message == null)
        {
            throw new ArgumentNullException(nameof(message));
        }

        var printDoc = new PrintDocument();
        printDoc.PrintPage += (s, e) =>
        {
            e.Graphics.DrawString(message.TextBody, SystemFonts.DefaultFont, Brushes.Black, e.MarginBounds.Left, e.MarginBounds.Top, StringFormat.GenericTypographic);
            e.HasMorePages = false;
        };
        printDoc.Print();
    }
}
