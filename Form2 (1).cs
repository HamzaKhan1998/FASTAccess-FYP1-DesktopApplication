using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;


namespace FAST_Access
{
    public partial class Form2 : Form
    {
        public string username = "Admin";
        public string password = "Password";
        public Form2()
        {
            /*            Thread myThread = new Thread(new ThreadStart(StartSplashScreen));
                        myThread.Start();
                        Thread.Sleep(3000);
                        myThread.Abort();
             */
         
            SetStyle(ControlStyles.SupportsTransparentBackColor, true);
            this.BackColor = Color.Transparent;
            InitializeComponent();
         

        }

        public void Form2_Load(object sender, EventArgs e)
        {
            // Set Form's Transperancy 100 %
            this.Opacity = 0;

            // Start the Timer To Animate Form
            timer1.Enabled = true;
            this.WindowState = FormWindowState.Normal;
            textBox2.PasswordChar = '*';
        
        }
        public void StartSplashScreen()
        {

            Application.Run(new Form1());
        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            
        }

        private void pictureBox1_Click(object sender, EventArgs e)
        {

        }

        private void pictureBox2_Click(object sender, EventArgs e)
        {

        }

        private void button_WOC1_Click(object sender, EventArgs e)
        {
            if (textBox1.Text == username && textBox2.Text == password)
            {

         
                Form3 F3 = new Form3(this);
                F3.Show();
                this.Hide();
            }
            else
            {
                MessageBox.Show("Incorrect Login Details");
            }
        }

        private void button_WOC2_Click(object sender, EventArgs e)
        {
            Application.Exit();            
        }

        
        private void button_WOC3_Click(object sender, EventArgs e)
        {
            WindowState = FormWindowState.Minimized;         
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            this.Opacity += 0.07;
        }
    }
}
