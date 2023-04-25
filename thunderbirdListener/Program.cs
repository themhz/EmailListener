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
using System.Text;

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
        string password = config["Password"]; //To Create application specific password go https://myaccount.google.com/apppasswords
        string imapServer = config["ImapServer"];
        int imapPort = int.Parse(config["ImapPort"]);

        Console.WriteLine($"2.Program started waiting for order...");
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

        Font printFont = new Font("Arial", 8); // Adjust the font size to fit the paper width

        // Use message.TextBody instead of the subject
        string[] lines = message.TextBody.Split(new[] { "\r\n", "\r", "\n" }, StringSplitOptions.None);
        List<string> wrappedLines = new List<string>();

        var printDoc = new PrintDocument();

        using (Graphics graphics = Graphics.FromHwnd(IntPtr.Zero))
        {
            foreach (string line in lines)
            {
                string wrappedText = WrapText(line, printFont, 295, graphics);
                wrappedLines.AddRange(wrappedText.Split(new[] { "\r\n", "\r", "\n" }, StringSplitOptions.None));
            }
        }

        int totalLines = wrappedLines.Count;
        float lineHeight = printFont.GetHeight();
        int paperHeight = (int)(totalLines * lineHeight) + 20; // Calculate paper height based on the content

        printDoc.DefaultPageSettings.PaperSize = new PaperSize("Custom", 315, paperHeight); // Set dynamic paper height
        printDoc.DefaultPageSettings.Margins = new Margins(10, 10, 10, 10);

        printDoc.PrintPage += (s, e) =>
        {
            float y = e.MarginBounds.Top;

            for (int i = 0; i < wrappedLines.Count; i++)
            {
                e.Graphics.DrawString(wrappedLines[i], printFont, Brushes.Black, e.MarginBounds.Left, y, StringFormat.GenericTypographic);
                y += lineHeight;
            }

            e.HasMorePages = false;
        };

        printDoc.Print();
    }



    private static string WrapText(string text, Font font, float maxWidth, Graphics graphics)
    {
        string[] words = text.Split(' ');
        StringBuilder wrappedText = new StringBuilder();
        StringBuilder currentLine = new StringBuilder();

        foreach (string word in words)
        {
            string testLine = currentLine.Length == 0 ? word : currentLine.ToString() + " " + word;
            float testLineWidth = graphics.MeasureString(testLine, font).Width;

            if (testLineWidth > maxWidth)
            {
                if (currentLine.Length > 0)
                {
                    wrappedText.AppendLine(currentLine.ToString());
                    currentLine.Clear();
                }
            }
            else
            {
                if (currentLine.Length > 0)
                {
                    currentLine.Append(" ");
                }
            }
            currentLine.Append(word);
        }

        if (currentLine.Length > 0)
        {
            wrappedText.AppendLine(currentLine.ToString());
        }

        return wrappedText.ToString();
    }

}
