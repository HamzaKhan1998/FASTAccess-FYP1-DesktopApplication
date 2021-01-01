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
    public partial class Form1 : Form
    {
        

        public Form1()
        {
         /*   Thread myThread = new Thread(new ThreadStart(StartSplashScreen));
            myThread.Start();
            Thread.Sleep(3000);
            myThread.Abort();*/
            InitializeComponent();
 //           System.Media.SoundPlayer player = new System.Media.SoundPlayer();
 //           player.SoundLocation = @"D:\MusicProjects\LQD\AdnanLQDThemeMusicIntro.wav";
           // player.Play();

        }

        private void Form1_Load(object sender,EventArgs e)
        {

        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            progressBar1.Maximum = 1000;
            timer1.Start();
            progressBar1.Increment(1);
            
            if (progressBar1.Value==1000)
            {
                System.Media.SoundPlayer player = new System.Media.SoundPlayer();
                timer1.Stop();
                Form2 F2 = new Form2();
              //  if (progressBar1.Value==500)
              //  {
               //     
                //    player.SoundLocation = @"D:\MusicProjects\LQD\AdnanLQDThemeMusicIntro.wav";
                 //   player.Play();
                  //  
               // }
               // player.Stop();
                this.Hide();
                F2.Show();
            }
        }

        private void pictureBox1_Click(object sender, EventArgs e)
        {

        }

        private void Form1_Load_1(object sender, EventArgs e)
        {

        }

        private void timer2_Tick(object sender, EventArgs e)
        {

        }

        private void progressBar1_Click(object sender, EventArgs e)
        {

        }
    }
}
