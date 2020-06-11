using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace FAST_Access
{
    
    public partial class Form5 : Form
    {
       /* bool enable_button = new bool();
        bool enable_button1 = new bool();
        Form2 originalForm;
        public Form5(Form2 incomingForm)
        {
            InitializeComponent();
            this.submit.Enabled = false;
            submit.ButtonColor = Color.Gray;
            submit.BorderColor = Color.Gray;
            this.pictureBox3.Hide();
            this.pictureBox4.Hide();
            this.hideYourProfile();
        //    System.Media.SoundPlayer player = new System.Media.SoundPlayer();
        //    player.SoundLocation = @"D:\Uni\FYP\WTFA.wav";
         //   player.Play();
        }

        private void pictureBox2_Click(object sender, EventArgs e)
        {

        }

        private void button_WOC1_Click(object sender, EventArgs e)
        {

        }




        private void button_WOC3_Click(object sender, EventArgs e)
        {
            WindowState = FormWindowState.Minimized;
        }

        private void button_WOC2_Click_1(object sender, EventArgs e)
        {
            Form2 f = new Form2();
            f.Show();
            this.Hide();
        }

        private void Form3_Load(object sender, EventArgs e)
        {

        }

        private void button_WOC5_Click(object sender, EventArgs e)
        {
            this.enable_button = true;
            
            if (enable_button == true && enable_button1 == true)
            {
                submit.Enabled = true;
                submit.ButtonColor = Color.RoyalBlue;
                submit.BorderColor = Color.RoyalBlue;
                submit.OnHoverButtonColor = Color.Aqua;
                submit.OnHoverBorderColor = Color.Aqua;
            }

        }

        private void pictureBox2_Click1(object sender, EventArgs e)
        {

        }

        private void button_WOC4_Click(object sender, EventArgs e)
        {
            this.enable_button1 = true;
            this.pictureBox4.Show();
            if (enable_button == true && enable_button1 == true)
            {
                submit.Enabled = true;
                submit.ButtonColor = Color.RoyalBlue;
                submit.BorderColor = Color.RoyalBlue;
                submit.OnHoverButtonColor = Color.Aqua;
                submit.OnHoverBorderColor = Color.Aqua;
            }


        }

        

        private void button_WOC5_Click_1(object sender, EventArgs e)
        {
            this.enable_button = true;
            this.pictureBox3.Show();
            if (enable_button == true && enable_button1 == true)
            {
                submit.Enabled = true;
                submit.ButtonColor = Color.RoyalBlue;
                submit.BorderColor = Color.RoyalBlue;
                submit.OnHoverButtonColor = Color.Aqua;
                submit.OnHoverBorderColor = Color.Aqua;
            }
        }

        private void button_WOC4_Click_1(object sender, EventArgs e)
        {
            this.enable_button1 = true;
            this.pictureBox4.Show();
            if (enable_button == true && enable_button1 == true)
            {
                submit.Enabled = true;
                submit.ButtonColor = Color.RoyalBlue;
                submit.BorderColor = Color.RoyalBlue;
                submit.OnHoverButtonColor = Color.Aqua;
                submit.OnHoverBorderColor = Color.Aqua;
            }
        }

        private void button_WOC1_Click_1(object sender, EventArgs e)
        {
            if (submit.Enabled == true)
            {
                MessageBox.Show("Please wait while the time-table is being generated");
                Task t = Task.Delay(3000);
                t.Wait();
                MessageBox.Show("Time-Table generated successfully with no constraint violations");
                Task t1 = Task.Delay(3000);
                t1.Wait();
                string fileExcel;
                fileExcel = @"C:\Users\Ubaid Qaiser\Downloads\Week4 Summer 2019.xlsx";
                Excel.Application xlapp;
                Excel.Workbook xlworkbook;
                xlapp = new Excel.Application();

                xlworkbook = xlapp.Workbooks.Open(fileExcel, 0, false, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);

                xlapp.Visible = true;
            }
        }

        private void pictureBox2_Click_1(object sender, EventArgs e)
        {

        }

        private void homeToolStripMenuItem_Click(object sender, EventArgs e)
        {
            showHome();
        }

        private void updateProfileToolStripMenuItem_Click(object sender, EventArgs e)
        {
            hideHome();


        }

        private void checkBox2_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void pictureBox3_Click(object sender, EventArgs e)
        {

        }



        private void hideHome()
        {
            submit.Hide();
            button_WOC2.Hide();
            button_WOC3.Hide();
            
            pictureBox3.Hide();
            pictureBox4.Hide();
        }
        private void showHome()
        {
            submit.Show();
            button_WOC2.Show();
            button_WOC3.Show();
            
            if (enable_button1==true)
            {
                this.pictureBox4.Show();
            }
            if (enable_button == true)
            {
                this.pictureBox3.Show();
            }
            if (enable_button == true && enable_button1 == true)
            {
                Button.submit.Enabled = true;
                submit.ButtonColor = Color.RoyalBlue;
                submit.BorderColor = Color.RoyalBlue;
                submit.OnHoverButtonColor = Color.Aqua;
                submit.OnHoverBorderColor = Color.Aqua;
            }

        }

        private void manageUsersToolStripMenuItem_Click(object sender, EventArgs e)
        {
            hideHome();
        }

        private void updateCoursesToolStripMenuItem_Click(object sender, EventArgs e)
        {
            hideHome();
        }

        private void updateTeachersToolStripMenuItem_Click(object sender, EventArgs e)
        {
            hideHome();
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }

        private void hideYourProfile()
        {
            textBox1.Hide();
            textBox2.Hide();
            textBox3.Hide();
            buttonSubmit.Hide();
        }

        private void buttonSubmit_Click(object sender, EventArgs e)
        {
            originalForm.username = textBox1.Text;
            originalForm.password = textBox2.Text;
        }*/
    }
}
