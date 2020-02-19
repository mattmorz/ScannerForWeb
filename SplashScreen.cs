using Microsoft.Win32;
using System;
using System.Windows.Forms;

namespace NewScan
{

    public partial class SplashScreen : Form
    {
        RegistryKey registry = Registry.CurrentUser.OpenSubKey(@"Software\Microsoft\Windows\CurrentVersion\Run", true);
        public SplashScreen()
        {
            InitializeComponent();
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            timer1.Start();
            progressBar1.Increment(1);
            if (progressBar1.Value == 100)
            {
                timer1.Stop();
                Form1 f1 = new Form1();
                f1.Show();
                this.Hide();

            }
        }

        private void pictureBox1_Click(object sender, EventArgs e)
        {

        }

        private void SplashScreen_Load(object sender, EventArgs e)
        {
            registry.SetValue("Scanner for Web", Application.ExecutablePath.ToString());
        }

        private void label1_Click(object sender, EventArgs e)
        {

        }
    }
}
