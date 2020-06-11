using System;
using System.Collections.Generic;
using System.Data.SQLite;
using System.Diagnostics;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;
using IronPython.Hosting;

namespace FAST_Access
{
    static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
        {
            //Console.WriteLine("Welcome to Python");
            //var psi = new ProcessStartInfo();
            //psi.FileName = @"C:\Program Files (x86)\Microsoft Visual Studio\Shared\Python37_64\python.exe";
            //var script = @"C:\Users\Hamza\.PyCharmCE2019.3\config\scratches\GA+DB4.py";  
            //psi.UseShellExecute = false;
            //psi.CreateNoWindow = true;
            //psi.RedirectStandardOutput = true;
            //var engine = Python.CreateEngine();
            //var script = @"C:\Users\Hamza\.PyCharmCE2019.3\config\scratches\GA+DB4.py";
            //var source = engine.CreateScriptSourceFromFile(script);
            //var scope = engine.CreateScope();
            //source.Execute(scope);
            /*
            Database dbobject = new Database();
            string query = "SELECT * from instructor";
            SQLiteCommand mycommand = new SQLiteCommand(query, dbobject.myConnection);
            dbobject.OpenConnection();
            SQLiteDataReader result = mycommand.ExecuteReader();
            if (result.HasRows)
            {
                while(result.Read())
                {
                    Console.WriteLine("ID: {0} - Name: {1}", result["InstructorID"], result["Name"]);
                }
            }

            dbobject.CloseConnection(); */
            

            //Console.ReadKey();

            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new Form1());
        }
    }
}
