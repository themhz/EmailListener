using TavernaLayorgosPrinterConsole;
using thunderbirdListener;
using static Org.BouncyCastle.Math.EC.ECCurve;

namespace TavernaLayorgosPrinter
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }        
        
        private async void button2_Click(object sender, EventArgs e)
        {
            MessageBox.Show("starting");
            await EmailPrinterProcessor.ReadEmailsAsync(EmailPrinterProcessor.LoadConfiguration());
        }

        private void Form1_FormClosed(object sender, FormClosedEventArgs e)
        {
            TokenCancelator.startProcess = 0;
        }

        private void button3_Click(object sender, EventArgs e)
        {
            TokenCancelator.startProcess = 0;
        }

        public static class TokenCancelator
        {
            public static int startProcess = 1;
        }
    }
}