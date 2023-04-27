using Microsoft.Extensions.Configuration;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using MailKit.Net.Imap;
using MailKit;
using MimeKit;
using System.Drawing.Printing;
using MimeKit;
using System.Drawing;
using MailKit.Search;
using Serilog;
using System.Diagnostics;


namespace thunderbirdListener
{
    public static class EmailPrinterProcessor
    {

        private static NotifyIcon trayIcon;
        private static ContextMenuStrip trayMenu;
        private static IConfiguration config;
        private static string logFilePath = "log.txt";
        private static LoggerConfiguration loger;

        public static IConfiguration LoadConfiguration()
        {
            var outputPath = AppDomain.CurrentDomain.BaseDirectory;
            var configPath = Path.Combine(outputPath, "config.ini");

            var builder = new ConfigurationBuilder()
                .SetBasePath(outputPath)
                .AddIniFile(configPath, optional: false, reloadOnChange: true);

            return builder.Build();
        }


        public static async Task ReadEmailsAsync(IConfiguration config, Panel pnlStatus, TextBox txtStatus)
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

            Log.Information("1.Program started waiting for order...");
            txtStatus.Text += "1.Program started waiting for order...\r\n";
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
                    pnlStatus.BackColor = Color.Green;
                    Log.Information("await Task.Delay(TimeSpan.FromSeconds(20));");
                    txtStatus.Text += "await Task.Delay(TimeSpan.FromSeconds(20));\r\n";
                    var query = SearchQuery.And(
                        SearchQuery.DeliveredAfter(lastProcessedMessageDate.DateTime),
                        SearchQuery.NotSeen
                    );

                    await client.NoOpAsync();
                    var newMessages = await inbox.SearchAsync(query);

                    Log.Information($"Reading messages {newMessages.Count}");
                    txtStatus.Text += $"Reading messages {newMessages.Count}\r\n";
                    foreach (var uid in newMessages)
                    {
                        var message = await inbox.GetMessageAsync(uid);

                        if (message.From.Mailboxes.FirstOrDefault().Address == config["MailFromDev"] && message.Subject.Contains("Comanda"))
                        {
                            Log.Information($"Message received at: {DateTime.Now.ToString("MM/dd/yyyy hh:mm tt")}");
                            txtStatus.Text += $"Message received at: {DateTime.Now.ToString("MM/dd/yyyy hh:mm tt")}\r\n";
                            Log.Information($"Subject: {message.Subject}");
                            txtStatus.Text += $"Subject: {message.Subject}\r\n";
                            Log.Information($"From: {message.From}");
                            txtStatus.Text += $"From: {message.From} \r\n";
                            Log.Information($"Date: {message.Date}");
                            txtStatus.Text += $"Date: {message.Date} \r\n";
                            Log.Information($"Body: {message.TextBody}");
                            txtStatus.Text += $"Body: {message.TextBody} \r\n";
                            Log.Information("-------------------------");
                            txtStatus.Text += "-------------------------\r\n";

                            txtStatus.Text += message.TextBody;
                            PrintMessageBody(message);

                            // Mark the message as read
                            await inbox.SetFlagsAsync(uid, MessageFlags.Seen, true);

                            // Update the last processed message date
                            lastProcessedMessageDate = message.Date;
                        }
                    }

                    pnlStatus.BackColor = Color.White;
                    // Wait for one minute before checking for new messages
                    //await Task.Delay(TimeSpan.FromMinutes(1));
                    await Task.Delay(TimeSpan.FromSeconds(20));
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
        public static string WrapText(string text, Font font, float maxWidth, Graphics graphics)
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
}
