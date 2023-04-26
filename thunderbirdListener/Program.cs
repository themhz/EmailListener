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
using MailKit.Search;
using Serilog;

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
        string basePath = AppDomain.CurrentDomain.BaseDirectory;        
        string logFilePath = Path.Combine(basePath, "log.txt");

        Log.Logger = new LoggerConfiguration()
        .MinimumLevel.Debug()
        .WriteTo.Console()
        .WriteTo.File(logFilePath)
        .CreateLogger();

        // Read email credentials and IMAP settings from the configuration
        string email = config["Email"];
        string password = config["Password"];
        string imapServer = config["ImapServer"];
        int imapPort = int.Parse(config["ImapPort"]);

        Console.WriteLine(".Program started waiting for order...");

        DateTimeOffset lastProcessedMessageDate = DateTimeOffset.Now.AddDays(-1);

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
            await inbox.OpenAsync(FolderAccess.ReadWrite);

            // Monitor the mailbox for new messages
            while (true)
            {
                // Wait for one minute before checking for new messages
                //await Task.Delay(TimeSpan.FromMinutes(1));
                await Task.Delay(TimeSpan.FromSeconds(30));

                // Search for messages received after the last processed message date
                var query = SearchQuery.And(
                    SearchQuery.DeliveredAfter(lastProcessedMessageDate.DateTime),
                    SearchQuery.NotSeen
                );

                var newMessages = await inbox.SearchAsync(query);

                foreach (var uid in newMessages)
                {
                    var message = await inbox.GetMessageAsync(uid);

                    if (message.From.Mailboxes.FirstOrDefault().Address == "themhz@gmail.com" && message.Subject.Contains("Comanda"))
                    {
                        Log.Information($"Message received at: {DateTime.Now.ToString("MM/dd/yyyy hh:mm tt")}");
                        Log.Information($"Subject: {message.Subject}");
                        Log.Information($"From: {message.From}");
                        Log.Information($"Date: {message.Date}");
                        Log.Information($"Body: {message.TextBody}");
                        Log.Information("-------------------------");

                        //PrintMessageBody(message);

                        // Mark the message as read
                        await inbox.SetFlagsAsync(uid, MessageFlags.Seen, true);

                        // Update the last processed message date
                        lastProcessedMessageDate = message.Date;
                    }
                }
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
