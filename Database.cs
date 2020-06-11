using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.SQLite;
using System.IO;

namespace FAST_Access
{
    class Database
    {
        public SQLiteConnection myConnection;

        public Database()
        {
            Console.WriteLine("Hello");
            
            
            if (!File.Exists("./class_schedule_4.db"))
            {
                Console.WriteLine("Innnnnnnnn");
                SQLiteConnection.CreateFile("class_schedule_4.db");
            }
        }

        public void OpenConnection()
        {
            if (myConnection.State != System.Data.ConnectionState.Open)
            {
                myConnection.Open();
            }
        }

        public void CloseConnection()
        {
            if (myConnection.State != System.Data.ConnectionState.Closed)
            {
                myConnection.Close();
            }
        }
    }
}
