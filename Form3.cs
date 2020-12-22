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
using IronPython.Hosting;
using IronPython.SQLite;
using Microsoft.Scripting;
using Microsoft.Scripting.Hosting;
using System.Data.SQLite;
using System.Data.SqlClient;
using System.Text.RegularExpressions;
using Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using DataTable = System.Data.DataTable;
using TextBox = System.Windows.Forms.TextBox;
using System.IO;
using OfficeOpenXml;
using ClosedXML.Excel;
using System.Collections;

namespace FAST_Access
{

    public partial class Form3 : Form
    {
        int rn = 312;
        bool enable_button = new bool();
        bool enable_button1 = new bool();
        BackgroundWorker bgw = new BackgroundWorker();
        Form2 originalForm = new Form2();
        public Form3(Form2 incomingForm)
        {
            Python.CreateEngine();
            originalForm = incomingForm;
            InitializeComponent();
            this.button_WOC1.Enabled = true;
            this.hideYourProfile();
            this.hideRooms();
            this.hideCourses();
            this.hideTeachers();
            FillCombobox();
            fillCoreElective();
            fillCourseLab();
            FillCombobox2();
//            FillCombobox1();
            fillDays();
            fillCourses();
            FillSection();
            //   fillCourseLab();
            this.textBox4.Enabled = false;
            this.textBox4.BackColor = Color.DarkGray;

            button_WOC1.ButtonColor = Color.RoyalBlue;
            button_WOC1.BorderColor = Color.RoyalBlue;
            button_WOC1.OnHoverButtonColor = Color.Aqua;
            button_WOC1.OnHoverBorderColor = Color.Aqua;


        }

        private SQLiteConnection sqlcon;
        private SQLiteCommand sqlcmd;
        private SQLiteDataAdapter DB;
        private SQLiteDataAdapter DB1;
        private SQLiteDataAdapter DB2;
        private SQLiteDataAdapter DB3;
        private SQLiteDataAdapter DB4;
        private DataSet DS = new DataSet();
        private DataSet DS1 = new DataSet();
        private DataSet DS2 = new DataSet();
        private DataSet DS3 = new DataSet();
        private DataSet DS4 = new DataSet();
        private DataTable DT = new DataTable();
        private DataTable DT1 = new DataTable();
        private DataTable DT2 = new DataTable();
        private DataTable DT3 = new DataTable();
        private DataTable DT4 = new DataTable();
        private IEnumerable<object> activities;

        public void SetConnection()
        {
            sqlcon = new SQLiteConnection("Data Source = class_schedule_4.db");
        }

        public void ExecuteQuery(string txtQuery)
        {
            SetConnection();
            sqlcon.Open();
            sqlcmd = sqlcon.CreateCommand();
            sqlcmd.CommandText = txtQuery;
            sqlcmd.ExecuteNonQuery();
            sqlcon.Close();

        }

        public void FillSection()
        {
            this.Section.Items.Add("1-A");
            this.Section.Items.Add("1-B");
            this.Section.Items.Add("1-C");
            this.Section.Items.Add("1-D");
            this.Section.Items.Add("1-E");
            this.Section.Items.Add("1-F");

            this.Section.Items.Add("2-A");
            this.Section.Items.Add("2-B");
            this.Section.Items.Add("2-C");
            this.Section.Items.Add("2-D");
            this.Section.Items.Add("2-E");
            this.Section.Items.Add("2-F");

            this.Section.Items.Add("3-A");
            this.Section.Items.Add("3-B");
            this.Section.Items.Add("3-C");
            this.Section.Items.Add("3-D");
            this.Section.Items.Add("3-E");
            this.Section.Items.Add("3-F");

            this.Section.Items.Add("4-A");
            this.Section.Items.Add("4-B");
            this.Section.Items.Add("4-C");
            this.Section.Items.Add("4-D");
            this.Section.Items.Add("4-E");
            this.Section.Items.Add("4-F");
        }

        protected void FillCombobox()
        {
            sqlcon = new SQLiteConnection
            ("Data Source=class_schedule_4.db;Version=3;new=False;Compress=True;");

            SetConnection();
            sqlcon.Open();
            sqlcmd = sqlcon.CreateCommand();
            string CommandText = "select RNumber from room";
            DB = new SQLiteDataAdapter(CommandText, sqlcon);
            DS.Reset();
            DB.Fill(DS);
            DT = DS.Tables[0];
            comboBox1.DisplayMember = "RNumber";
            comboBox1.ValueMember = "RNumber";
            comboBox1.DataSource = DS.Tables[0];
            sqlcon.Close();

        }

        protected void FillCombobox2()
        {
            sqlcon = new SQLiteConnection
            ("Data Source=class_schedule_4.db;Version=3;new=False;Compress=True;");

            SetConnection();
            sqlcon.Open();
            sqlcmd = sqlcon.CreateCommand();
            string CommandText = "select CourseName from course";
            DB = new SQLiteDataAdapter(CommandText, sqlcon);
            DS.Reset();
            DB.Fill(DS);
            DT = DS.Tables[0];
            comboBox2.DisplayMember = "CourseName";
            comboBox2.ValueMember = "CourseName";
            comboBox2.DataSource = DS.Tables[0];
            sqlcon.Close();

        }

        protected void FillCombobox1()
        {
            sqlcon = new SQLiteConnection
            ("Data Source=class_schedule_4.db;Version=3;new=False;Compress=True;");

            SetConnection();
            sqlcon.Open();
            sqlcmd = sqlcon.CreateCommand();
            string CommandText = "select Name from instructor";
            DB = new SQLiteDataAdapter(CommandText, sqlcon);
            DS.Reset();
            DB.Fill(DS);
            DT = DS.Tables[0];
            teacherID.DataSource = DS.Tables[0];
            teacherID.BindingContext = new BindingContext();
            teacherID.DisplayMember = "Name";
            teacherID.ValueMember = "Name";

            sqlcon.Close();


        } 

        public void fillCourseLab()
        {
            this.classLab.Items.Add("Class");
            this.classLab.Items.Add("Lab");
        }
        public void fillCoreElective()
        {
            this.coreElective.Items.Add("Core");
            this.coreElective.Items.Add("Elective");
            this.coreElective.Items.Add("Core Lab");
            this.coreElective.Items.Add("Elective Lab");

        }
        public void fillDays()
        {

            this.comboBox6.Items.Add("Consecutive");
            this.comboBox6.Items.Add("Non Consecutive");
            this.comboBox9.Items.Add("Consecutive");
            this.comboBox9.Items.Add("Non Consecutive");

            //------------------------------------------------------------------------------------------------

            //1hourx3
            this.P3C1.Items.Add("Monday, Wednesday, Friday");//1
            this.P3C1.Items.Add("Monday, Thursday, Friday");//5
            this.P3C1.Items.Add("Tuesday, Wednesday, Thursday");//2
            this.P3C1.Items.Add("Wednesday, Thursday, Friday");//3
            this.P3C1.Items.Add("Tuesday, Thursday, Friday");//4
            this.P3C1.Items.Add("Monday, Tuesday, Wednesday");//6
            //1.5x2
            this.P3C1.Items.Add("Monday, Tuesday");//1
            this.P3C1.Items.Add("Wednesday, Thursday");//5
            this.P3C1.Items.Add("Tuesday, Wednesday");//2
            this.P3C1.Items.Add("Monday, Friday");//4
            this.P3C1.Items.Add("Tuesday, Thursday");//3

            //3x1
            this.P3C1.Items.Add("Monday");
            this.P3C1.Items.Add("Tuesday");
            this.P3C1.Items.Add("Wednesday");
            this.P3C1.Items.Add("Thursday");
            this.P3C1.Items.Add("Friday");

            //------------------------------------------------------------------------------------------------

            //1hourx3
            this.P3C2.Items.Add("Monday, Wednesday, Friday");//1
            this.P3C2.Items.Add("Monday, Thursday, Friday");//5
            this.P3C2.Items.Add("Tuesday, Wednesday, Thursday");//2
            this.P3C2.Items.Add("Wednesday, Thursday, Friday");//3
            this.P3C2.Items.Add("Tuesday, Thursday, Friday");//4
            this.P3C2.Items.Add("Monday, Tuesday, Wednesday");//6
            //1.5x2
            this.P3C2.Items.Add("Monday, Tuesday");//1
            this.P3C2.Items.Add("Wednesday, Thursday");//5
            this.P3C2.Items.Add("Tuesday, Wednesday");//2
            this.P3C2.Items.Add("Monday, Friday");//4
            this.P3C2.Items.Add("Tuesday, Thursday");//3


            //3x1
            this.P3C2.Items.Add("Monday");
            this.P3C2.Items.Add("Tuesday");
            this.P3C2.Items.Add("Wednesday");
            this.P3C2.Items.Add("Thursday");
            this.P3C2.Items.Add("Friday");
            

            //------------------------------------------------------------------------------------------------


            //1hourx3
            this.P3C3.Items.Add("Monday, Wednesday, Friday");//1
            this.P3C3.Items.Add("Monday, Thursday, Friday");//5
            this.P3C3.Items.Add("Tuesday, Wednesday, Thursday");//2
            this.P3C3.Items.Add("Wednesday, Thursday, Friday");//3
            this.P3C3.Items.Add("Tuesday, Thursday, Friday");//4
            this.P3C3.Items.Add("Monday, Tuesday, Wednesday");//6

            //1.5x2
            this.P3C3.Items.Add("Monday, Tuesday");//1
            this.P3C3.Items.Add("Wednesday, Thursday");//5
            this.P3C3.Items.Add("Tuesday, Wednesday");//2
            this.P3C3.Items.Add("Monday, Friday");//4
            this.P3C3.Items.Add("Tuesday, Thursday");//3

            //3x1
            this.P3C3.Items.Add("Monday");
            this.P3C3.Items.Add("Tuesday");
            this.P3C3.Items.Add("Wednesday");
            this.P3C3.Items.Add("Thursday");
            this.P3C3.Items.Add("Friday");

            //------------------------------------------------------------------------------------------------

            //1hourx3
            this.comboBox10.Items.Add("Monday, Wednesday, Friday");//1
            this.comboBox10.Items.Add("Monday, Thursday, Friday");//5
            this.comboBox10.Items.Add("Tuesday, Wednesday, Thursday");//2
            this.comboBox10.Items.Add("Wednesday, Thursday, Friday");//3
            this.comboBox10.Items.Add("Tuesday, Thursday, Friday");//4
            this.comboBox10.Items.Add("Monday, Tuesday, Wednesday");//6
            //1.5x2
            this.comboBox10.Items.Add("Monday, Tuesday");//1
            this.comboBox10.Items.Add("Wednesday, Thursday");//5
            this.comboBox10.Items.Add("Tuesday, Wednesday");//2
            this.comboBox10.Items.Add("Monday, Friday");//4
            this.comboBox10.Items.Add("Tuesday, Thursday");//3

            //3x1
            this.comboBox10.Items.Add("Monday");
            this.comboBox10.Items.Add("Tuesday");
            this.comboBox10.Items.Add("Wednesday");
            this.comboBox10.Items.Add("Thursday");
            this.comboBox10.Items.Add("Friday");



            //------------------------------------------------------------------------------------------------

            this.P1C1.Items.Add("3hr x 1 class");
            this.P1C1.Items.Add("2hr x 1 class");
            this.P1C1.Items.Add("1.5hr x 2 class");
            this.P1C1.Items.Add("1hr x 3 class");
            this.P1C1.Items.Add("2hr x 1 + 1hr x 1 class"); ///////////////new

            this.P1C2.Items.Add("3hr x 1 class");
            this.P1C2.Items.Add("2hr x 1 class");
            this.P1C2.Items.Add("1.5hr x 2 class");
            this.P1C2.Items.Add("1hr x 3 class");
            this.P1C2.Items.Add("2hr x 1 + 1hr x 1 class"); ///////////////new

            this.P1C3.Items.Add("3hr x 1 class");
            this.P1C3.Items.Add("2hr x 1 class");
            this.P1C3.Items.Add("1.5hr x 2 class");
            this.P1C3.Items.Add("1hr x 3 class");
            this.P1C3.Items.Add("2hr x 1 + 1hr x 1 class"); ///////////////new

            this.comboBox12.Items.Add("3hr x 1 class");
            this.comboBox12.Items.Add("2hr x 1 class");
            this.comboBox12.Items.Add("1.5hr x 2 class");
            this.comboBox12.Items.Add("1hr x 3 class");
            this.comboBox12.Items.Add("2hr x 1 + 1hr x 1 class"); ///////////////new

            //------------------------------------------------------------------------------------------------

            this.P2C1.Items.Add("First Half");
            this.P2C1.Items.Add("Second Half");

            this.P2C2.Items.Add("First Half");
            this.P2C2.Items.Add("Second Half");

            this.P2C3.Items.Add("First Half");
            this.P2C3.Items.Add("Second Half");

            this.comboBox11.Items.Add("First Half");
            this.comboBox11.Items.Add("Second Half");



        }
        public void fillCourses()
        {
            sqlcon = new SQLiteConnection("Data Source=class_schedule_4.db;Version=3;new=False;Compress=True;");
            sqlcon.Open();
            sqlcmd = sqlcon.CreateCommand();
            string CommandText = "select CourseName from course";
            DB = new SQLiteDataAdapter(CommandText, sqlcon);
            DS.Reset();
            DB.Fill(DS);
            DT = DS.Tables[0];
            coursebox3.DataSource = DS.Tables[0];
            coursebox3.BindingContext = new BindingContext();
            coursebox3.DisplayMember = "CourseName";
            coursebox3.ValueMember = "CourseName";

            coursecomboBox4.DataSource = DS.Tables[0];
            coursecomboBox4.BindingContext = new BindingContext();
            coursecomboBox4.DisplayMember = "CourseName";
            coursecomboBox4.ValueMember = "CourseName";

            coursecomboBox5.DataSource = DS.Tables[0];
            coursecomboBox5.BindingContext = new BindingContext();
            coursecomboBox5.DisplayMember = "CourseName";
            coursecomboBox5.ValueMember = "CourseName";
/*
            coursecomboBox13.DataSource = DS.Tables[0];
            coursecomboBox13.BindingContext = new BindingContext();
            coursecomboBox13.DisplayMember = "CourseName";
            coursecomboBox13.ValueMember = "CourseName";
            */

            string CommandText2 = "select Name from instructor";
            DB2 = new SQLiteDataAdapter(CommandText2, sqlcon);
            DS2.Reset();
            DB2.Fill(DS2);

            teacherID.DataSource = DS2.Tables[0];
            teacherID.BindingContext = new BindingContext();
            teacherID.DisplayMember = "Name";
            teacherID.ValueMember = "Name";

            comboBox7.DataSource = DS2.Tables[0];
            comboBox7.BindingContext = new BindingContext();
            comboBox7.DisplayMember = "Name";
            comboBox7.ValueMember = "Name";


            string CommandText3 = "select RNumber from room";
            DB3 = new SQLiteDataAdapter(CommandText3, sqlcon);
            DS3.Reset();
            DB3.Fill(DS3);

            comboBox3.DataSource = DS3.Tables[0];
            comboBox3.BindingContext = new BindingContext();
            comboBox3.DisplayMember = "RNumber";
            comboBox3.ValueMember = "RNumber";

            comboBox4.DataSource = DS3.Tables[0];
            comboBox4.BindingContext = new BindingContext();
            comboBox4.DisplayMember = "RNumber";
            comboBox4.ValueMember = "RNumber";

            comboBox5.DataSource = DS3.Tables[0];
            comboBox5.BindingContext = new BindingContext();
            comboBox5.DisplayMember = "RNumber";
            comboBox5.ValueMember = "RNumber";

            comboBox8.DataSource = DS3.Tables[0];
            comboBox8.BindingContext = new BindingContext();
            comboBox8.DisplayMember = "RNumber";
            comboBox8.ValueMember = "RNumber";

            sqlcon.Close();

        }
        public void fillCourses1()
        {
            sqlcon = new SQLiteConnection("Data Source=class_schedule_4.db;Version=3;new=False;Compress=True;");
            sqlcon.Open();
            sqlcmd = sqlcon.CreateCommand();
            string CommandText = "select CourseName from course";
            DB = new SQLiteDataAdapter(CommandText, sqlcon);
            DS.Reset();
            DB.Fill(DS);
            DT = DS.Tables[0];
            coursebox3.DataSource = DS.Tables[0];
            coursebox3.BindingContext = new BindingContext();
            coursebox3.DisplayMember = "CourseName";
            coursebox3.ValueMember = "CourseName";
            sqlcon.Close();

        }
        public void fillCourses2()
        {
            sqlcon = new SQLiteConnection("Data Source=class_schedule_4.db;Version=3;new=False;Compress=True;");

            SetConnection();
            sqlcon.Open();
            sqlcmd = sqlcon.CreateCommand();
            string CommandText = "select CourseName from course";
            DB = new SQLiteDataAdapter(CommandText, sqlcon);
            DS.Reset();
            DB.Fill(DS);
            DT = DS.Tables[0];
            coursecomboBox4.DataSource = DS.Tables[0];
            coursecomboBox4.BindingContext = new BindingContext();
            coursecomboBox4.DisplayMember = "CourseName";
            coursecomboBox4.ValueMember = "CourseName";

            sqlcon.Close();
        }
        public void fillCourses3()
        {
            sqlcon = new SQLiteConnection("Data Source=class_schedule_4.db;Version=3;new=False;Compress=True;");

            SetConnection();
            sqlcon.Open();
            sqlcmd = sqlcon.CreateCommand();
            string CommandText = "select CourseName from course";
            DB = new SQLiteDataAdapter(CommandText, sqlcon);
            DS.Reset();
            DB.Fill(DS);
            DT = DS.Tables[0];
            coursecomboBox5.DataSource = DS.Tables[0];
            coursecomboBox5.BindingContext = new BindingContext();
            coursecomboBox5.DisplayMember = "CourseName";
            coursecomboBox5.ValueMember = "CourseName";
            sqlcon.Close();

        }
        private void loaddata()
        {
            
        }
        private void pictureBox2_Click(object sender, EventArgs e)
        {
            
        }
        private void button_WOC1_Click(object sender, EventArgs e)
        {

            if (button_WOC1.Enabled==true)
            {
                ScriptEngine engine = Python.CreateEngine();
                engine.ExecuteFile(@"C:\Users\Ubaid Qaiser\Desktop\tr.py");
                MessageBox.Show("Please wait while the time-table is being generated");
                Task t = Task.Delay(3000);
                t.Wait();

                MessageBox.Show("Time-Table generated successfully with no constraint violations");
                Task t1=Task.Delay(3000);
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
        private void button_WOC3_Click(object sender, EventArgs e)
        {
            WindowState = FormWindowState.Minimized;
        }

        private void button_WOC2_Click_1(object sender, EventArgs e)
        {
            originalForm.Show();
            this.Hide();
        }

        private void Form3_Load(object sender, EventArgs e)
        {
            
        }


        private void pictureBox2_Click1(object sender, EventArgs e)
        {

        }

        

        

        void bgw_DoWork(object sender, DoWorkEventArgs e)
        {
            int total = 57; //some number (this is your variable to change)!!

            for (int i = 0; i <= total; i++) //some number (total)
            {
                System.Threading.Thread.Sleep(100);
                int percents = (i * 100) / total;
                bgw.ReportProgress(percents, i);
                //2 arguments:
                //1. procenteges (from 0 t0 100) - i do a calcumation 
                //2. some current value!
            }
        }

        void bgw_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            progressBar1.Value = e.ProgressPercentage;
        }

        void bgw_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            //do the code when bgv completes its work
        }


        private void button_WOC1_Click_1(object sender, EventArgs e)
        {


            progressBar1.Show();
            var py = Python.CreateEngine();
            ICollection<string> searchPaths = py.GetSearchPaths();
            searchPaths.Add("C:\\Users\\ubaid\\PycharmProjects\\pythonProject\\venv\\Lib\\site-packages\\prettytable");
            py.SetSearchPaths(searchPaths);
            py.ExecuteFile("C:\\Users\\ubaid\\PycharmProjects\\pythonProject\\main.py");
            MessageBox.Show("Please wait while the time-table is being generated");
            bgw.DoWork += new DoWorkEventHandler(bgw_DoWork);
            bgw.ProgressChanged += new ProgressChangedEventHandler(bgw_ProgressChanged);
            bgw.RunWorkerCompleted += new RunWorkerCompletedEventHandler(bgw_RunWorkerCompleted);
            bgw.WorkerReportsProgress = true;
            bgw.RunWorkerAsync();

        }

            private void pictureBox2_Click_1(object sender, EventArgs e)
        {

        }


        private void homeToolStripMenuItem_Click(object sender, EventArgs e)
        {
            showHome();
            hideYourProfile();
            hideRooms();
            hideCourses();
            hideTeachers();
        }

        private void updateProfileToolStripMenuItem_Click(object sender, EventArgs e)
        {
            hideHome();
            showYourProfile();
            hideRooms();
            hideCourses();
            hideTeachers();
        }

        private void checkBox2_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void pictureBox3_Click(object sender, EventArgs e)
        {

        }

        private void hideHome()
        {
            button_WOC1.Hide();
            progressBar1.Hide();
            button_WOC5.Hide();
            button_WOC6.Hide();
            button_WOC4.Hide();
            button_WOC7.Hide();
            button_WOC8.Hide();
            this.dataGridView1.Hide();
            button_WOC11.Hide();
            Select_Excel.Hide();

        }
        private void showHome()
        {
            progressBar1.Hide();
            button_WOC1.Show();
            button_WOC5.Show();
            button_WOC6.Show();
            button_WOC4.Show();
            button_WOC7.Show();
            button_WOC8.Show();
            this.dataGridView1.Show();
            button_WOC11.Show();
            Select_Excel.Show();

        }

        private void showYourProfile()
        {
            pictureBox5.Show();
            pictureBox6.Show();
            textBox1.Show();
            textBox2.Show();
            textBox3.Show();
            buttonSubmit.Show();
            this.dataGridView1.Hide();
        }
        private void hideYourProfile()
        {
            pictureBox5.Hide();
            pictureBox6.Hide();
            textBox1.Hide();
            textBox2.Hide();
            textBox3.Hide();
            buttonSubmit.Hide();
        }

        private void manageUsersToolStripMenuItem_Click(object sender, EventArgs e)
        {
            hideHome();
            hideTeachers();
            hideCourses();
            hideRooms();
            showYourProfile();
            this.dataGridView1.Hide();
        }

        private void updateCoursesToolStripMenuItem_Click(object sender, EventArgs e)
        {
            hideTeachers();
            hideHome();
            hideRooms();
            hideYourProfile();
            showCourses();
            this.dataGridView1.Hide();
            FillCombobox2();
            
        }

        private void updateTeachersToolStripMenuItem_Click(object sender, EventArgs e)
        {
            hideHome();
            hideCourses();
            hideRooms();
            hideYourProfile();
            showTeachers();
            this.dataGridView1.Hide();
 //           FillCombobox1();
            fillCourses();
            fillTcourse();
            

        }
        


        private void buttonSubmit_Click(object sender, EventArgs e)
        {
            this.originalForm.username = textBox1.Text;
            this.originalForm.password = textBox2.Text;
            MessageBox.Show("Password and username updated successfully");
        }
        private void hideRooms()
        {
            this.textBox4.Hide();
            this.comboBox1.Hide();
            this.AddRoom.Hide();
            this.Remove.Hide();
            this.ViewRooms.Hide();
            this.classLab.Hide();
            this.dataGridView4.Hide();
            this.dataGridView2.Hide();
            this.dataGridView1.Hide();
            this.dataGridView3.Hide();
        }
        private void showRooms()
        {
            this.textBox4.Show();
            this.comboBox1.Show();
            this.AddRoom.Show();
            this.Remove.Show();
            this.ViewRooms.Show();
            this.classLab.Show();
        }
        private void textBox4_TextChanged(object sender, EventArgs e)
        {

        }
        private void updateRoomsToolStripMenuItem_Click(object sender, EventArgs e)
        {
            hideYourProfile();
            hideHome();
            showRooms();
            hideCourses();
            hideTeachers();
            this.dataGridView1.Hide();
            FillCombobox();
        }
        private void hideCourses()
        {
            this.Section.Hide();
            this.courseName.Hide();
//            this.sectionsnumericUpDown1.Hide();
            this.comboBox2.Hide();
//            this.courseCode.Hide();
            this.AddCourse.Hide();
//            this.coursecomboBox2.Hide();
 //           this.courseName2.Hide();
 //           this.sectionsnumericUpDown2.Hide();
 //           this.coursecomboBox2.Hide();
            this.RemoveCourse.Hide();
 //           this.updateCourse.Hide();
            this.AddCourse.Hide();
            this.ViewCourses.Hide();
            this.dataGridView4.Hide();
            this.dataGridView2.Hide();
            this.dataGridView1.Hide();
            this.dataGridView3.Hide();
            this.RemoveCourse.Hide();
            this.ViewCourses.Hide();
 //           this.Day.Hide();
            this.coreElective.Hide();
 //           this.Timing.Hide();
        }
        private void showCourses()
        {
            this.Section.Show();
            this.courseName.Show();
//            this.sectionsnumericUpDown1.Show();
//            this.courseCode.Show();
            this.comboBox2.Show();
            this.AddCourse.Show();
//            this.coursecomboBox2.Show();
//            this.courseName2.Show();
//            this.sectionsnumericUpDown2.Show();
//            this.coursecomboBox2.Show();
            this.RemoveCourse.Show();
//            this.updateCourse.Show();
            this.AddCourse.Show();
            this.ViewCourses.Show();
//            this.Day.Show();
            this.coreElective.Show();
//            this.Timing.Show();
            this.RemoveCourse.Show();
            this.ViewCourses.Show();
            this.dataGridView1.Hide();
        }
        private void hideTeachers()
        {
            this.teacherName.Hide();
            this.teacherID.Hide();
            this.coursebox3.Hide();
            this.coursecomboBox4.Hide();
            this.coursecomboBox5.Hide();
            this.comboBox3.Hide();
            this.comboBox4.Hide();
            this.comboBox5.Hide();
            this.comboBox6.Hide();
            this.AddTeacher.Hide();
//            this.updateTeacher.Hide();
            this.RemoveTeacher.Hide();
            this.ViewTeachers.Hide();
            this.P1C1.Hide();
            this.P1C2.Hide();
            this.P1C3.Hide();
            this.P2C1.Hide();
            this.P2C2.Hide();
            this.P2C3.Hide();
            this.P3C1.Hide();
            this.P3C2.Hide();
            this.P3C3.Hide();
            this.dataGridView4.Hide();
            this.dataGridView2.Hide();
            this.dataGridView1.Hide();
            this.dataGridView3.Hide();

            this.comboBox7.Hide();
            this.comboBox8.Hide();
            this.comboBox9.Hide();
            this.comboBox10.Hide();
            this.comboBox11.Hide();
            this.comboBox12.Hide();
            this.coursecomboBox13.Hide();
            this.button_WOC9.Hide();
            this.button_WOC10.Hide();
        }
        private void showTeachers()
        {
            this.teacherName.Show();
            this.teacherID.Show();
            this.coursebox3.Show();
            this.coursecomboBox4.Show();
            this.coursecomboBox5.Show();
            this.comboBox3.Show();
            this.comboBox4.Show();
            this.comboBox5.Show();
            this.comboBox6.Show();
            this.AddTeacher.Show();
            this.RemoveTeacher.Show();
            this.ViewTeachers.Show();
            this.dataGridView1.Hide();
            this.P1C1.Show();
            this.P1C2.Show();
            this.P1C3.Show();
            this.P2C1.Show();
            this.P2C2.Show();
            this.P2C3.Show();
            this.P3C1.Show();
            this.P3C2.Show();
            this.P3C3.Show();
            this.comboBox7.Show();
            this.comboBox8.Show();
            this.comboBox9.Show();
            this.comboBox10.Show();
            this.comboBox11.Show();
            this.comboBox12.Show();
            this.coursecomboBox13.Show();
            this.button_WOC9.Show();
            this.button_WOC10.Show();
        }

        private static readonly Regex _regex = new Regex("[^0-9]+"); //regex that matches disallowed text
        

        private void AddRoom_Click(object sender, EventArgs e)
        {
            bool IsTextAllowed(string text)
            {
                return !_regex.IsMatch(text);
            }
            SetConnection();
            sqlcon.Open();
            sqlcmd = sqlcon.CreateCommand();
            string s = textBox4.Text;
            string Check = classLab.SelectedItem.ToString();
            var textBox = sender as TextBox;
            if (Check == "Lab")
            {
                string CommandText = "Insert into room values('" + s + "','Lab',50)";
                DB = new SQLiteDataAdapter(CommandText, sqlcon);
                DS.Reset();
                DB.Fill(DS);
                sqlcon.Close();
            }
            else if (Check == "Class" && IsTextAllowed(textBox4.Text)==true)
            {
                string CommandText = "Insert into room values('" +"C"+ s + "','Classroom', 50)";
                DB = new SQLiteDataAdapter(CommandText, sqlcon);
                DS.Reset();
                DB.Fill(DS);
                sqlcon.Close();
            }

        }



        private void sectionnumericUpDown7_ValueChanged(object sender, EventArgs e)
        {

        }

        public void ViewRooms_Click(object sender, EventArgs e)
        {

            SetConnection();
            sqlcon.Open();
            sqlcmd = sqlcon.CreateCommand();
            string CommandText = "select * from room";
            DB = new SQLiteDataAdapter(CommandText, sqlcon);
            DS.Reset();
            DB.Fill(DS);
            DT = DS.Tables[0];
            dataGridView4.DataSource = DT;
            sqlcon.Close();
            this.dataGridView4.Show();

        }

        private void ViewTeachers_Click(object sender, EventArgs e)
        {

            SetConnection();
            sqlcon.Open();
            sqlcmd = sqlcon.CreateCommand();
            string CommandText = "select InstructorID, Name from instructor";
            DB = new SQLiteDataAdapter(CommandText, sqlcon);
            DS.Reset();
            DB.Fill(DS);
            DT = DS.Tables[0];
            dataGridView3.DataSource = DT;
            sqlcon.Close();
            this.dataGridView3.Show();

            //dbobject.CloseConnection();
        }

        private void ViewCourses_Click(object sender, EventArgs e)
        {
            FillCombobox2();
            SetConnection();
            sqlcon.Open();
            sqlcmd = sqlcon.CreateCommand();
            string CommandText = "select CourseID, CourseName, CourseType, Section,Teacher from course";
            DB = new SQLiteDataAdapter(CommandText, sqlcon);
            DS.Reset();
            DB.Fill(DS);
            DT = DS.Tables[0];
            dataGridView2.DataSource = DT;
            sqlcon.Close();
            this.dataGridView2.Show();


        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            
        }

        private void Remove_Click(object sender, EventArgs e)
        {
            SetConnection();
            sqlcon.Open();
            sqlcmd = sqlcon.CreateCommand();
            string s = comboBox1.Text;
            string CommandText = "delete from room WHERE RNumber='"+ s +"'";
            DB = new SQLiteDataAdapter(CommandText, sqlcon);
            DS.Reset();
            DB.Fill(DS);
            sqlcon.Close();
            ViewRooms_Click(sender, e);
            FillCombobox();

        }

        private void RemoveCourse_Click(object sender, EventArgs e)
        {
            SetConnection();
            sqlcon.Open();
            sqlcmd = sqlcon.CreateCommand();
            string s1 = comboBox2.Text;
            string CommandText = "delete from course WHERE CourseName='" + s1 + "'";
            DB = new SQLiteDataAdapter(CommandText, sqlcon);
            DS.Reset();
            DB.Fill(DS);
            sqlcon.Close();
            ViewCourses_Click(sender, e);
            FillCombobox2();
        }

        private void RemoveTeacher_Click(object sender, EventArgs e)
        {
            SetConnection();
            sqlcon.Open();
            sqlcmd = sqlcon.CreateCommand();
            string s = teacherID.Text;
            string CommandText = "delete from instructor WHERE Name='" + s + "'";
            DB = new SQLiteDataAdapter(CommandText, sqlcon);
            DS.Reset();
            DB.Fill(DS);
            sqlcon.Close();
            ViewTeachers_Click(sender, e);
            FillCombobox1();
        }

        private void teacherID_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void coursebox3_SelectedIndexChanged(object sender, EventArgs e)
        {
            string c1 = coursebox3.Text;
            if (c1.Contains("LAB"))
            {
                P1C1.Text = "3hr x 1 class";
                P1C1.Enabled = false;
                P1C1.BackColor = Color.DarkGray;
            }
            else
            {
                P1C1.Enabled = true;
                P1C1.BackColor = Color.WhiteSmoke;
            }
        }

        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void classLab_SelectedIndexChanged(object sender, EventArgs e)
        {
            string Check = classLab.SelectedItem.ToString();
            if (Check == "Lab")
            {
                textBox4.Enabled = true;
                textBox4.BackColor = Color.WhiteSmoke;
            }
            else
            {

                textBox4.Enabled = true;
                textBox4.BackColor = Color.WhiteSmoke;
            }
        }

        private void AddCourse_Click(object sender, EventArgs e)
        {
            bool IsTextAllowed(string text)
            {
                return !_regex.IsMatch(text);
            }
            SetConnection();
            sqlcon.Open();
            sqlcmd = sqlcon.CreateCommand();
            string s = courseName.Text;
            string sec = Section.Text;
            string Check = coreElective.SelectedItem.ToString();
            var textBox = sender as TextBox;
            if (Check == "Core" && IsTextAllowed(courseName.Text) == false &&sec != "Section")
            {
                s = s + " (" + sec + ")";
                string CommandText = "Insert into course(coursename,coursetype,section,maxnumofstudents,tpreference,dpreference,qpreference,conflictwith, conflictvalue, rpreference, concount) values('" + s + "','Core','"+sec+ "',45,1,0,0, 'N/A', 'N/A', 'N/A', 0)";
                DB = new SQLiteDataAdapter(CommandText, sqlcon);
                DS.Reset();
                DB.Fill(DS);
                
            }
            else if (Check == "Elective" && IsTextAllowed(courseName.Text) == false && sec != "Section")
            {
                s = s + " (" + sec + ")";
                string CommandText = "Insert into course(coursename,coursetype,section,maxnumofstudents,tpreference,dpreference,qpreference,conflictwith, conflictvalue, rpreference, concount) values('" + s + "','Elective','" + sec + "',45,1.5,0,0, 'N/A', 'N/A', 'N/A', 0)";
                DB = new SQLiteDataAdapter(CommandText, sqlcon);
                DS.Reset();
                DB.Fill(DS);
                
            }
            else if (Check == "Elective Lab" && IsTextAllowed(courseName.Text) == false && sec != "Section")
            {
                s = s + " LAB " + " (" + sec + ")";
                string CommandText = "Insert into course(coursename,coursetype,section,maxnumofstudents,tpreference,dpreference,qpreference,conflictwith, conflictvalue, rpreference, concount) values('" + s + "','Elective Lab','" + sec + "',45,3,0,0, 'N/A', 'N/A', 'N/A' ,0)";
                DB = new SQLiteDataAdapter(CommandText, sqlcon);
                DS.Reset();
                DB.Fill(DS);
                
            }
            else if (Check == "Core Lab" && IsTextAllowed(courseName.Text) == false && sec != "Section")
            {
                s = s + " LAB " + " (" + sec + ")";
                string CommandText = "Insert into course(coursename,coursetype,section,maxnumofstudents,tpreference,dpreference,qpreference,conflictwith, conflictvalue, rpreference, concount) values('" + s + "','Core Lab','" + sec + "',45,3,0,0, 'N/A', 'N/A', 'N/A', 0)";
                DB = new SQLiteDataAdapter(CommandText, sqlcon);
                DS.Reset();
                DB.Fill(DS);
                
            }
            string s1 = QueryResult("select courseid from course where coursename = '" + s + "' ");
            int x1 = Int32.Parse(s1);
            string CommandText7 = string.Format("Insert into dept_course (name, course_number) values ('CS', {0})", x1);
            DB1 = new SQLiteDataAdapter(CommandText7, sqlcon);
            DS1.Reset();
            DB1.Fill(DS1);
            sqlcon.Close();

        }

        private void Section_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        public string QueryResult(string query)
        {
            string result = "";
            SQLiteConnection sqlite = new SQLiteConnection("Data Source = class_schedule_4.db");
            try
            {
                sqlite.Open();  //Initiate connection to the db
                SQLiteCommand cmd = sqlite.CreateCommand();
                cmd.CommandText = query;  //set the passed query
                result = cmd.ExecuteScalar().ToString();
            }
            finally
            {
                sqlite.Close();
            }
            return result;
        }

        private void AddTeacher_Click(object sender, EventArgs e)
        {
            bool IsTextAllowed(string text)
            {
                return !_regex.IsMatch(text);
            }
            SetConnection();
            sqlcon.Open();
            sqlcmd = sqlcon.CreateCommand();
            string s = teacherName.Text;
//            string Check = coreElective.SelectedItem.ToString();
            var textBox = sender as TextBox;
            string c1 = coursebox3.Text;
            string c2 = coursecomboBox4.Text;
            string c3 = coursecomboBox5.Text;

            string r1 = comboBox3.Text;
            string r2 = comboBox4.Text;
            string r3 = comboBox5.Text;
            string d1 = P1C1.Text;
            string d2 = P1C2.Text;
            string d3 = P1C3.Text;
            string t1 = P2C1.Text;
            string t2 = P2C2.Text;
            string t3 = P2C3.Text;
            string day1 = P1C3.Text;
            string day2 = P2C3.Text;
            string day3 = P3C3.Text;
            string dd1,dd2,dd3;
            int dd11;
            int dd22;
            int dd33;
            double qq11;
            double qq22;
            double qq33;
            int tt11=0;
            int tt22=0;
            int tt33=0;
            int consec = 0;

            if (IsTextAllowed(teacherName.Text) == false)
            {

                if (comboBox6.Text == "Consecutive")
                {
                    consec = 1;
                }
                else
                {
                    consec = 0;
                }

                string CommandText2 = "Insert into instructor(name,spacer,spacercheck) values('" + s + "',"+consec+",0)";
                DB2 = new SQLiteDataAdapter(CommandText2, sqlcon);
                DS2.Reset();
                DB2.Fill(DS2);

                



                if (P1C1.Text == "3hr x 1 class")
                {

                    P2C1.Enabled = false;
                    P2C1.BackColor = Color.DarkGray;

                    qq11 = 3;
                    dd1 = P3C1.Text;
                    if (dd1 == "Monday")
                    {
                        dd11 = 1;
                    }
                    else if (dd1 == "Tuesday")
                    {
                        dd11 = 2;
                    }
                    else if (dd1 == "Wednesday")
                    {
                        dd11 = 3;
                    }
                    else if (dd1 == "Thursday")
                    {
                        dd11 = 4;
                    }
                    else
                    {
                        dd11 = 5;
                    }

                }
                else if (P1C1.Text == "2hr x 1 class")
                {
                    P2C1.Enabled = true;
                    P2C1.BackColor = Color.WhiteSmoke;
                    qq11 = 2;
                    dd1 = P3C1.Text;
                    if (dd1 == "Monday")
                    {
                        dd11 = 1;
                    }
                    else if (dd1 == "Tuesday")
                    {
                        dd11 = 2;
                    }
                    else if (dd1 == "Wednesday")
                    {
                        dd11 = 3;
                    }
                    else if (dd1 == "Thursday")
                    {
                        dd11 = 4;
                    }
                    else
                    {
                        dd11 = 5;
                    }

                }
                else if (P1C1.Text == "1.5hr x 2 class")
                {
                    P2C1.Enabled = true;
                    P2C1.BackColor = Color.WhiteSmoke;
                    qq11 = 1.5;
                    dd1 = P3C1.Text;
                    if (dd1 == "Monday, Tuesday")
                    {
                        dd11 = 1;
                    }
                    else if (dd1 == "Tuesday, Wednesday")
                    {
                        dd11 = 2;
                    }
                    else if (dd1 == "Tuesday, Thursday")
                    {
                        dd11 = 3;
                    }
                    else if (dd1 == "Monday, Friday")
                    {
                        dd11 = 4;
                    }
                    else
                    {
                        dd11 = 5;
                    }

                }
                else 
                {
                    P2C1.Enabled = true;
                    P2C1.BackColor = Color.WhiteSmoke;
                    qq11 = 1;
                    dd1 = P3C1.Text;
                    if (dd1 == "Monday, Wednesday, Friday")
                    {
                        dd11 = 1;
                    }
                    else if (dd1 == "Tuesday, Wednesday, Thursday")
                    {
                        dd11 = 2;
                    }
                    else if (dd1 == "Wednesday, Thursday, Friday")
                    {
                        dd11 = 3;
                    }
                    else if (dd1 == "Tuesday, Thursday, Friday")
                    {
                        dd11 = 4;
                    }
                    else
                    {
                        dd11 = 5;
                    }

                }


                if (P1C2.Text == "3hr x 1 class")
                {

                    P2C2.Enabled = false;
                    P2C2.BackColor = Color.DarkGray;


                    qq22 = 3;
                    dd1 = P3C2.Text;
                    if (dd1 == "Monday")
                    {
                        dd22 = 1;
                    }
                    else if (dd1 == "Tuesday")
                    {
                        dd22 = 2;
                    }
                    else if (dd1 == "Wednesday")
                    {
                        dd22 = 3;
                    }
                    else if (dd1 == "Thursday")
                    {
                        dd22 = 4;
                    }
                    else
                    {
                        dd22 = 5;
                    }

                }
                else if (P1C2.Text == "2hr x 1 class")
                {
                    P2C2.Enabled = true;
                    P2C2.BackColor = Color.WhiteSmoke;
                    qq22 = 2;
                    dd1 = P3C2.Text;
                    if (dd1 == "Monday")
                    {
                        dd22 = 1;
                    }
                    else if (dd1 == "Tuesday")
                    {
                        dd22 = 2;
                    }
                    else if (dd1 == "Wednesday")
                    {
                        dd22 = 3;
                    }
                    else if (dd1 == "Thursday")
                    {
                        dd22 = 4;
                    }
                    else
                    {
                        dd22 = 5;
                    }

                }
                else if (P1C2.Text == "1.5hr x 2 class")
                {
                    P2C2.Enabled = true;
                    P2C2.BackColor = Color.WhiteSmoke;
                    qq22 = 1.5;
                    dd1 = P3C2.Text;
                    if (dd1 == "Monday, Tuesday")
                    {
                        dd22 = 1;
                    }
                    else if (dd1 == "Tuesday, Wednesday")
                    {
                        dd22 = 2;
                    }
                    else if (dd1 == "Tuesday, Thursday")
                    {
                        dd22 = 3;
                    }
                    else if (dd1 == "Monday, Friday")
                    {
                        dd22 = 4;
                    }
                    else
                    {
                        dd22 = 5;
                    }

                }
                else 
                {
                    P2C2.Enabled = true;
                    P2C2.BackColor = Color.WhiteSmoke;
                    qq22 = 1;
                    dd1 = P3C2.Text;
                    if (dd1 == "Monday, Wednesday, Friday")
                    {
                        dd22 = 1;
                    }
                    else if (dd1 == "Tuesday, Wednesday, Thursday")
                    {
                        dd22 = 2;
                    }
                    else if (dd1 == "Wednesday, Thursday, Friday")
                    {
                        dd22 = 3;
                    }
                    else if (dd1 == "Tuesday, Thursday, Friday")
                    {
                        dd22 = 4;
                    }
                    else
                    {
                        dd22 = 5;
                    }

                }



                if (P1C3.Text == "3hr x 1 class")
                {


                    qq33 = 3;
                    dd1 = P3C3.Text;
                    if (dd1 == "Monday")
                    {
                        dd33 = 1;
                    }
                    else if (dd1 == "Tuesday")
                    {
                        dd33 = 2;
                    }
                    else if (dd1 == "Wednesday")
                    {
                        dd33 = 3;
                    }
                    else if (dd1 == "Thursday")
                    {
                        dd33 = 4;
                    }
                    else
                    {
                        dd33 = 5;
                    }

                }
                else if (P1C3.Text == "2hr x 1 class")
                {
                    P2C3.Enabled = true;
                    P2C3.BackColor = Color.WhiteSmoke;
                    qq33 = 2;
                    dd1 = P3C3.Text;
                    if (dd1 == "Monday")
                    {
                        dd33 = 1;
                    }
                    else if (dd1 == "Tuesday")
                    {
                        dd33 = 2;
                    }
                    else if (dd1 == "Wednesday")
                    {
                        dd33 = 3;
                    }
                    else if (dd1 == "Thursday")
                    {
                        dd33 = 4;
                    }
                    else
                    {
                        dd33 = 5;
                    }

                }
                else if (P1C3.Text == "1.5hr x 2 class")
                {
                    P2C3.Enabled = true;
                    P2C3.BackColor = Color.WhiteSmoke;
                    qq33 = 1.5;
                    dd1 = P3C3.Text;
                    if (dd1 == "Monday, Tuesday")
                    {
                        dd33 = 1;
                    }
                    else if (dd1 == "Tuesday, Wednesday")
                    {
                        dd33 = 2;
                    }
                    else if (dd1 == "Tuesday, Thursday")
                    {
                        dd33 = 3;
                    }
                    else if (dd1 == "Monday, Friday")
                    {
                        dd33 = 4;
                    }
                    else
                    {
                        dd33 = 5;
                    }

                }
                else 
                {
                    P2C3.Enabled = true;
                    P2C3.BackColor = Color.WhiteSmoke;
                    qq33 = 1;
                    dd1 = P3C3.Text;
                    if (dd1 == "Monday, Wednesday, Friday")
                    {
                        dd33 = 1;
                    }
                    else if (dd1 == "Tuesday, Wednesday, Thursday")
                    {
                        dd33 = 2;
                    }
                    else if (dd1 == "Wednesday, Thursday, Friday")
                    {
                        dd33 = 3;
                    }
                    else if (dd1 == "Tuesday, Thursday, Friday")
                    {
                        dd33 = 4;
                    }
                    else
                    {
                        dd33 = 5;
                    }

                }


                if (P2C1.Text=="First Half")
                {
                    tt11 = 1;
                }
                else 
                {
                    tt11 = 2;
                }

                if (P2C2.Text == "First Half")
                {
                    tt22 = 1;
                }
                else
                {
                    tt22 = 2;
                }

                if (P2C3.Text == "First Half")
                {
                    tt33 = 1;
                }
                else 
                {
                    tt33 = 2;
                }
                //r1,r2,r3 consec

                string s1 = QueryResult("select courseid from course where coursename = '" + c1 + "' ");
                int x1 = Int32.Parse(s1);
                string s2 = QueryResult("select courseid from course where coursename = '" + c2 + "' ");
                int x2 = Int32.Parse(s2);
                string s3 = QueryResult("select courseid from course where coursename = '" + c3 + "' ");
                int x3 = Int32.Parse(s3);
                string r = QueryResult("select instructorid from instructor where name = '" + s + "' ");
                int x = Int32.Parse(r);
                string CommandText7 = string.Format("Insert into course_instructor (course_number, instructor_number) values ({0},{1})", x,x1);
                DB1 = new SQLiteDataAdapter(CommandText7, sqlcon);
                DS1.Reset();
                DB1.Fill(DS1);
                string CommandText5 = string.Format("Insert into course_instructor (course_number, instructor_number) values ({0},{1})", x, x2);
                DB1 = new SQLiteDataAdapter(CommandText5, sqlcon);
                DS1.Reset();
                DB1.Fill(DS1);
                string CommandText6 = string.Format("Insert into course_instructor (course_number, instructor_number) values ({0},{1})", x, x3);
                DB1 = new SQLiteDataAdapter(CommandText6, sqlcon);
                DS1.Reset();
                DB1.Fill(DS1);

                string CommandText = string.Format("Update course set Tpreference = {0} , Dpreference = {1} ,Qpreference = {2},Rpreference = '" + r1 + "',Teacher = '" + s + "' where coursename =  '" + c1 + "'",qq11,dd11,tt11);
                string CommandText4 = string.Format("Update course set Tpreference = {0} , Dpreference = {1} ,Qpreference = {2},Rpreference = '" + r2 + "',Teacher = '" + s + "' where coursename =  '" + c2 + "'", qq22, dd22, tt22);
                string CommandText3 = string.Format("Update course set Tpreference = {0} , Dpreference = {1} ,Qpreference = {2},Rpreference = '"+r3+ "',Teacher = '" + s + "' where coursename =  '" + c3 + "'", qq33, dd33, tt33);
                
                DB = new SQLiteDataAdapter(CommandText, sqlcon);
                DS.Reset();
                DB.Fill(DS);
                DB1 = new SQLiteDataAdapter(CommandText4, sqlcon);
                DS1.Reset();
                DB1.Fill(DS1);
                DB3 = new SQLiteDataAdapter(CommandText3, sqlcon);
                DS3.Reset();
                DB3.Fill(DS3);

                sqlcon.Close();
            }
            else
            {
                //Display invalid name message
            }


        }

        private void P1C1_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (P1C1.Text == "3hr x 1 class")
            {
                P2C1.Enabled = false;
                P2C1.BackColor = Color.DarkGray;
            }

            else
            {
                P2C1.Enabled = true;
                P2C1.BackColor = Color.WhiteSmoke;
            }

            if (P1C1.Text == "Duration")
            {
                //3x1
                this.P3C1.Items.Clear();
                this.P3C1.Enabled = false;
                this.P3C1.BackColor = Color.DarkGray;
            }
            if (P1C1.Text== "3hr x 1 class")
            {
                
                //3x1
                this.P3C1.Enabled = true;
                this.P3C1.BackColor = Color.WhiteSmoke;
                this.P3C1.Items.Clear();
                this.P3C1.Items.Add("Monday");
                this.P3C1.Items.Add("Tuesday");
                this.P3C1.Items.Add("Wednesday");
                this.P3C1.Items.Add("Thursday");
                this.P3C1.Items.Add("Friday");
            }
            else if(P1C1.Text == "2hr x 1 class")
            {
                
                this.P3C1.Enabled = true;
                this.P3C1.BackColor = Color.WhiteSmoke;
                this.P3C1.Items.Clear();
                this.P3C1.Items.Add("Monday");
                this.P3C1.Items.Add("Tuesday");
                this.P3C1.Items.Add("Wednesday");
                this.P3C1.Items.Add("Thursday");
                this.P3C1.Items.Add("Friday");
            }
            else if (P1C1.Text == "1.5hr x 2 class")
            {
                Console.WriteLine("1.5hr x 1 class");
                this.P3C1.Enabled = true;
                this.P3C1.BackColor = Color.WhiteSmoke;
                this.P3C1.Items.Clear();
                //1.5x2
                this.P3C1.Items.Add("Monday, Tuesday");//1
                this.P3C1.Items.Add("Wednesday, Thursday");//5
                this.P3C1.Items.Add("Tuesday, Wednesday");//2
                this.P3C1.Items.Add("Monday, Friday");//4
                this.P3C1.Items.Add("Tuesday, Thursday");//3
            }
            else if (P1C1.Text == "1hr x 3 class")
            {
                Console.WriteLine("1hr x 1 class");
                this.P3C1.Enabled = true;
                this.P3C1.BackColor = Color.WhiteSmoke;
                this.P3C1.Items.Clear();
                //1hourx3
                this.P3C1.Items.Add("Monday, Wednesday, Friday");//1
                this.P3C1.Items.Add("Monday, Thursday, Friday");//5
                this.P3C1.Items.Add("Tuesday, Wednesday, Thursday");//2
                this.P3C1.Items.Add("Wednesday, Thursday, Friday");//3
                this.P3C1.Items.Add("Tuesday, Thursday, Friday");//4
            }
            
        }

        private void P1C2_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (P1C2.Text == "3hr x 1 class")
            {
                P2C2.Enabled = false;
                P2C2.BackColor = Color.DarkGray;
            }

            else
            {
                P2C2.Enabled = true;
                P2C2.BackColor = Color.WhiteSmoke;
            }
            if (P1C2.Text == "Duration")
            {
                //3x1
                this.P3C2.Items.Clear();
                this.P3C2.Enabled = false;
                this.P3C2.BackColor = Color.DarkGray;
            }
            if (P1C2.Text == "3hr x 1 class")
            {
                //3x1
                this.P3C2.Enabled = true;
                this.P3C2.BackColor = Color.WhiteSmoke;
                this.P3C2.Items.Clear();
                this.P3C2.Items.Add("Monday");
                this.P3C2.Items.Add("Tuesday");
                this.P3C2.Items.Add("Wednesday");
                this.P3C2.Items.Add("Thursday");
                this.P3C2.Items.Add("Friday");
            }
            else if (P1C2.Text == "2hr x 1 class")
            {
                //3x1
                this.P3C2.Enabled = true;
                this.P3C2.BackColor = Color.WhiteSmoke;
                this.P3C2.Items.Clear();
                this.P3C2.Items.Add("Monday");
                this.P3C2.Items.Add("Tuesday");
                this.P3C2.Items.Add("Wednesday");
                this.P3C2.Items.Add("Thursday");
                this.P3C2.Items.Add("Friday");
            }
            else if (P1C2.Text == "1.5hr x 2 class")
            {
                //3x1
                this.P3C2.Enabled = true;
                this.P3C2.BackColor = Color.WhiteSmoke;
                this.P3C2.Items.Clear();

                //1.5x2
                this.P3C2.Items.Add("Monday, Tuesday");//1
                this.P3C2.Items.Add("Wednesday, Thursday");//5
                this.P3C2.Items.Add("Tuesday, Wednesday");//2
                this.P3C2.Items.Add("Monday, Friday");//4
                this.P3C2.Items.Add("Tuesday, Thursday");//3
            }
            else if (P1C2.Text == "1hr x 3 class")
            {
                //3x1
                this.P3C2.Enabled = true;
                this.P3C2.BackColor = Color.WhiteSmoke;
                this.P3C2.Items.Clear();
                //1hourx3
                this.P3C2.Items.Add("Monday, Wednesday, Friday");//1
                this.P3C2.Items.Add("Monday, Thursday, Friday");//5
                this.P3C2.Items.Add("Tuesday, Wednesday, Thursday");//2
                this.P3C2.Items.Add("Wednesday, Thursday, Friday");//3
                this.P3C2.Items.Add("Tuesday, Thursday, Friday");//4
            }

        }

        private void P1C3_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (P1C3.Text == "3hr x 1 class")
            {
                P2C3.Enabled = false;
                P2C3.BackColor = Color.DarkGray;
            }

            else
            {
                P2C3.Enabled = true;
                P2C3.BackColor = Color.WhiteSmoke;
            }
            if (P1C3.Text == "Duration")
            {
                //3x1
                this.P3C3.Items.Clear();
                this.P3C3.Enabled = false;
                this.P3C3.BackColor = Color.DarkGray;
            }
            if (P1C3.Text == "3hr x 1 class")
            {
                //3x1
                this.P3C3.Enabled = true;
                this.P3C3.BackColor = Color.WhiteSmoke;
                this.P3C3.Items.Clear();
                this.P3C3.Items.Add("Monday");
                this.P3C3.Items.Add("Tuesday");
                this.P3C3.Items.Add("Wednesday");
                this.P3C3.Items.Add("Thursday");
                this.P3C3.Items.Add("Friday");
            }
            else if (P1C3.Text == "2hr x 1 class")
            {
                //3x1
                this.P3C3.Enabled = true;
                this.P3C3.BackColor = Color.WhiteSmoke;
                this.P3C3.Items.Clear();
                this.P3C3.Items.Add("Monday");
                this.P3C3.Items.Add("Tuesday");
                this.P3C3.Items.Add("Wednesday");
                this.P3C3.Items.Add("Thursday");
                this.P3C3.Items.Add("Friday");
            }
            else if (P1C3.Text == "1.5hr x 2 class")
            {
                this.P3C3.Enabled = true;
                this.P3C3.BackColor = Color.WhiteSmoke;
                //3x1
                this.P3C3.Items.Clear();

                //1.5x2
                this.P3C3.Items.Add("Monday, Tuesday");//1
                this.P3C3.Items.Add("Wednesday, Thursday");//5
                this.P3C3.Items.Add("Tuesday, Wednesday");//2
                this.P3C3.Items.Add("Monday, Friday");//4
                this.P3C3.Items.Add("Tuesday, Thursday");//3
            }
            else if (P1C3.Text == "1hr x 3 class")
            {
                this.P3C3.Enabled = true;
                this.P3C3.BackColor = Color.WhiteSmoke;
                //3x1
                this.P3C3.Items.Clear();
                //1hourx3
                this.P3C3.Items.Add("Monday, Wednesday, Friday");//1
                this.P3C3.Items.Add("Monday, Thursday, Friday");//5
                this.P3C3.Items.Add("Tuesday, Wednesday, Thursday");//2
                this.P3C3.Items.Add("Wednesday, Thursday, Friday");//3
                this.P3C3.Items.Add("Tuesday, Thursday, Friday");//4
            }
        }

        private void button_WOC4_Click(object sender, EventArgs e)
        {
            SetConnection();
            sqlcon.Open();
            sqlcmd = sqlcon.CreateCommand();
            string CommandText = "select coursename,timing,teacher,room from course_timing where timing LIKE 'Mon%'";
            DB = new SQLiteDataAdapter(CommandText, sqlcon);
            DS.Reset();
            DB.Fill(DS);
            DT = DS.Tables[0];
            dataGridView1.DataSource = DT;
            sqlcon.Close();
            this.dataGridView1.Show();
        }

        private void button_WOC5_Click(object sender, EventArgs e)
        {
            SetConnection();
            sqlcon.Open();
            sqlcmd = sqlcon.CreateCommand();
            string CommandText = "select coursename,timing,teacher,room from course_timing where timing LIKE '%Tue%'";
            DB = new SQLiteDataAdapter(CommandText, sqlcon);
            DS.Reset();
            DB.Fill(DS);
            DT = DS.Tables[0];
            dataGridView1.DataSource = DT;
            sqlcon.Close();
            this.dataGridView1.Show();
        }

        private void button_WOC6_Click(object sender, EventArgs e)
        {
            SetConnection();
            sqlcon.Open();
            sqlcmd = sqlcon.CreateCommand();
            string CommandText = "select coursename,timing,teacher,room from course_timing where timing LIKE '%Wed%'";
            DB = new SQLiteDataAdapter(CommandText, sqlcon);
            DS.Reset();
            DB.Fill(DS);
            DT = DS.Tables[0];
            dataGridView1.DataSource = DT;
            sqlcon.Close();
            this.dataGridView1.Show();
        }

        private void button_WOC7_Click(object sender, EventArgs e)
        {
            SetConnection();
            sqlcon.Open();
            sqlcmd = sqlcon.CreateCommand();
            string CommandText = "select coursename,timing,teacher,room from course_timing where timing LIKE '%Thurs%'";
            DB = new SQLiteDataAdapter(CommandText, sqlcon);
            DS.Reset();
            DB.Fill(DS);
            DT = DS.Tables[0];
            dataGridView1.DataSource = DT;
            sqlcon.Close();
            this.dataGridView1.Show();
        }

        private void button_WOC8_Click(object sender, EventArgs e)
        {
            SetConnection();
            sqlcon.Open();
            sqlcmd = sqlcon.CreateCommand();
            string CommandText = "select coursename,timing,teacher,room from course_timing where timing LIKE '%Fri%'";
            DB = new SQLiteDataAdapter(CommandText, sqlcon);
            DS.Reset();
            DB.Fill(DS);
            DT = DS.Tables[0];
            dataGridView1.DataSource = DT;
            sqlcon.Close();
            this.dataGridView1.Show();
        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void coursecomboBox5_SelectedIndexChanged(object sender, EventArgs e)
        {
            string c3 = coursecomboBox5.Text;
            if (c3.Contains("LAB"))
            {
                P1C3.Text = "3hr x 1 class";
                P1C3.Enabled = false;
                P1C3.BackColor = Color.DarkGray;
            }
            else
            {
                P1C3.Enabled = true;
                P1C3.BackColor = Color.WhiteSmoke;
            }
            
        }

        private void coursecomboBox4_SelectedIndexChanged(object sender, EventArgs e)
        {
            string c2 = coursecomboBox4.Text;
            if (c2.Contains("LAB"))
            {
                P1C2.Text = "3hr x 1 class";
                P1C2.Enabled = false;
                P1C2.BackColor = Color.DarkGray;
            }
            else
            {
                P1C2.Enabled = true;
                P1C2.BackColor = Color.WhiteSmoke;
            }
        }

        private void button_WOC9_Click(object sender, EventArgs e)
        {

        }

        private void comboBox12_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (comboBox12.Text == "3hr x 1 class")
            {
                comboBox11.Enabled = false;
                comboBox11.BackColor = Color.DarkGray;
            }

            else
            {
                comboBox11.Enabled = true;
                comboBox11.BackColor = Color.WhiteSmoke;
            }
            if (comboBox12.Text == "Duration")
            {
                //3x1
                this.comboBox10.Items.Clear();
                this.comboBox10.Enabled = false;
                this.comboBox10.BackColor = Color.DarkGray;
            }
            if (comboBox12.Text == "3hr x 1 class")
            {
                //3x1
                this.comboBox10.Enabled = true;
                this.comboBox10.BackColor = Color.WhiteSmoke;
                this.comboBox10.Items.Clear();
                this.comboBox10.Items.Add("Monday");
                this.comboBox10.Items.Add("Tuesday");
                this.comboBox10.Items.Add("Wednesday");
                this.comboBox10.Items.Add("Thursday");
                this.comboBox10.Items.Add("Friday");
            }
            else if (comboBox12.Text == "2hr x 1 class")
            {
                //3x1
                this.comboBox10.Enabled = true;
                this.comboBox10.BackColor = Color.WhiteSmoke;
                this.comboBox10.Items.Clear();
                this.comboBox10.Items.Add("Monday");
                this.comboBox10.Items.Add("Tuesday");
                this.comboBox10.Items.Add("Wednesday");
                this.comboBox10.Items.Add("Thursday");
                this.comboBox10.Items.Add("Friday");
            }
            else if (comboBox12.Text == "1.5hr x 2 class")
            {
                this.comboBox10.Enabled = true;
                this.comboBox10.BackColor = Color.WhiteSmoke;
                //3x1
                this.comboBox10.Items.Clear();

                //1.5x2
                this.comboBox10.Items.Add("Monday, Tuesday");//1
                this.comboBox10.Items.Add("Wednesday, Thursday");//5
                this.comboBox10.Items.Add("Tuesday, Wednesday");//2
                this.comboBox10.Items.Add("Monday, Friday");//4
                this.comboBox10.Items.Add("Tuesday, Thursday");//3
            }
            else if (comboBox12.Text == "1hr x 3 class")
            {
                this.comboBox10.Enabled = true;
                this.comboBox10.BackColor = Color.WhiteSmoke;
                //3x1
                this.P3C3.Items.Clear();
                //1hourx3
                this.comboBox10.Items.Add("Monday, Wednesday, Friday");//1
                this.comboBox10.Items.Add("Monday, Thursday, Friday");//5
                this.comboBox10.Items.Add("Tuesday, Wednesday, Thursday");//2
                this.comboBox10.Items.Add("Wednesday, Thursday, Friday");//3
                this.comboBox10.Items.Add("Tuesday, Thursday, Friday");//4
            }
        }

        private void coursecomboBox13_SelectedIndexChanged(object sender, EventArgs e)
        {
            string c3 = coursecomboBox13.Text;
            if (c3.Contains("LAB"))
            {
                comboBox12.Text = "3hr x 1 class";
                comboBox12.Enabled = false;
                comboBox12.BackColor = Color.DarkGray;
            }
            else
            {
                comboBox12.Enabled = true;
                comboBox12.BackColor = Color.WhiteSmoke;
            }
            sqlcon = new SQLiteConnection("Data Source=class_schedule_4.db;Version=3;new=False;Compress=True;");
            sqlcon.Open();
            sqlcmd = sqlcon.CreateCommand();

            string tname = comboBox7.Text;
            string CommandText = "select coursename from course WHERE teacher = '" + tname + "'";
            Console.WriteLine("Hello", tname);
            DB = new SQLiteDataAdapter(CommandText, sqlcon);
            DS.Reset();
            DB.Fill(DS);

            coursecomboBox13.DataSource = DS.Tables[0];
            coursecomboBox13.BindingContext = new BindingContext();
            coursecomboBox13.DisplayMember = "CourseName";
            coursecomboBox13.ValueMember = "CourseName";


            sqlcon.Close();

        }

        void fillTcourse()
        {


        }

        private void comboBox7_SelectedIndexChanged(object sender, EventArgs e)
        {


        }

        private void button2_Click(object sender, EventArgs e)
        {

        }

        private void button_WOC12_Click(object sender, EventArgs e)
        {
            string excelpath = @"C:\Users\ubaid\Desktop\ch.xlsx";
            Microsoft.Office.Interop.Excel.Application application = new Microsoft.Office.Interop.Excel.Application();
            Workbook workbook = application.Workbooks.Open(excelpath);

            for(int i = 1; i<= workbook.Sheets.Count; i++)
            {
                Worksheet worksheet = workbook.Worksheets[i];
                int c = 0;
                int j = 5;
                int inf = 0;
                while (c != 100)
                {
                    string cellvalue = ((Microsoft.Office.Interop.Excel.Range)worksheet.Cells[j, 3]).Value;
                    string cellvaluebelow = ((Microsoft.Office.Interop.Excel.Range)worksheet.Cells[j+1, 3]).Value;
                    MessageBox.Show(cellvalue);
                    if (cellvaluebelow == null)
                    {
                        MessageBox.Show("Null Value");
                        string sectionvalue = ((Microsoft.Office.Interop.Excel.Range)worksheet.Cells[j, 5]).Value;
                        string teachervalue = ((Microsoft.Office.Interop.Excel.Range)worksheet.Cells[j, 6]).Value;
                        string CommandText2 = "Insert into course (CourseName, Section, Teacher) values ('" + cellvalue + "', '" + sectionvalue + "', '" + teachervalue + "' )";
                        DB1 = new SQLiteDataAdapter(CommandText2, sqlcon);
                        DS1.Reset();
                        DB1.Fill(DS1);
                        j++;
                        while(cellvaluebelow == null)
                        {
                            if(inf == 10)
                            {
                                inf = 0;
                                break;
                            }
                            string sectionvalue1 = ((Microsoft.Office.Interop.Excel.Range)worksheet.Cells[j, 5]).Value;
                            string teachervalue1 = ((Microsoft.Office.Interop.Excel.Range)worksheet.Cells[j, 6]).Value;
                            string CommandText21 = "Insert into course (CourseName, Section, Teacher) values ('" + cellvalue + "', '" + sectionvalue1 + "', '" + teachervalue1 + "' )";
                            DB1 = new SQLiteDataAdapter(CommandText21, sqlcon);
                            DS1.Reset();
                            DB1.Fill(DS1);
                            j++;
                            cellvaluebelow = ((Microsoft.Office.Interop.Excel.Range)worksheet.Cells[j + 1, 3]).Value;
                            inf++;
                        }
                    }
                    else
                    {
                        string sectionvalue = ((Microsoft.Office.Interop.Excel.Range)worksheet.Cells[j, 5]).Value;
                        string teachervalue = ((Microsoft.Office.Interop.Excel.Range)worksheet.Cells[j, 6]).Value;
                        string CommandText2 = "Insert into course (CourseName, Section, Teacher) values ('" + cellvalue + "', '" + sectionvalue + "', '" + teachervalue + "' )";
                        //MessageBox.Show(CommandText2);
                        DB1 = new SQLiteDataAdapter(CommandText2, sqlcon);
                        DS1.Reset();
                        DB1.Fill(DS1);
                        j++;
                    }
                    //listBox1.Items.Add(cellvalue);
                    //MessageBox.Show(c);
                    c++;
                }
                
            }

            workbook.Close(false, excelpath, null);
            Marshal.ReleaseComObject(workbook);

        }

        private void listBox1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }



        private void button_WOC11_Click(object sender, EventArgs e)
        {
            sqlcon = new SQLiteConnection("Data Source=class_schedule_4.db;Version=3;new=False;Compress=True;");
            sqlcon.Open();
            
            using (sqlcmd = new SQLiteCommand("SELECT * FROM course_timing"))
            {
                using (SQLiteDataAdapter sqlda = new SQLiteDataAdapter())
                {
                    sqlcmd.Connection = sqlcon;
                    sqlda.SelectCommand = sqlcmd;

                    using (DataTable dt = new DataTable())
                    {
                        sqlda.Fill(dt);

                        using (XLWorkbook wb = new XLWorkbook())
                        {
                            string fname = "Desktop/TimeTable.xlsx";
                            wb.Worksheets.Add(dt, "course_timing");
                            wb.SaveAs(fname);
                            MessageBox.Show("Time Table Exported Successfuly");
                        }
                    }
                }
            }

        }

        private void button_WOC12_Click_1(object sender, EventArgs e)
        {
            progressBar1.Show();
            var py = Python.CreateEngine();
            ICollection<string> searchPaths = py.GetSearchPaths();
            searchPaths.Add("C:\\Users\\ubaid\\Desktop\\FYP\\venv\\Lib\\site-packages");
            py.SetSearchPaths(searchPaths);
            py.ExecuteFile("C:\\Users\\ubaid\\Desktop\\FYP");
            MessageBox.Show("Please wait while the time-table is being uploaded");
            bgw.DoWork += new DoWorkEventHandler(bgw_DoWork);
            bgw.ProgressChanged += new ProgressChangedEventHandler(bgw_ProgressChanged);
            bgw.RunWorkerCompleted += new RunWorkerCompletedEventHandler(bgw_RunWorkerCompleted);
            bgw.WorkerReportsProgress = true;
            bgw.RunWorkerAsync();
        }
    }
}
