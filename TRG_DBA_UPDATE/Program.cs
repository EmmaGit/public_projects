using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.SqlClient;
using System.Configuration;

namespace TRG_DBA_UPDATE
{
    public class Logger
    {
        public static void Memorize(string content) 
        {
            string filename = Properties.Settings.Default.MemoFile;
            try 
            {
                System.IO.File.AppendAllText(filename, content, Encoding.UTF8);
            }
            catch(Exception exception){ }
        }

        public static void Log(Exception exception = null) 
        {
            string filename = Properties.Settings.Default.LogFile;
            string content = DateTime.Now.ToString() + ";";
            if(exception == null)
            {
                exception = new Exception("No message");
            }
            try
            {
                content += "\"" + exception.Message + ";\"" + exception.Source + ";\r\n";
                System.IO.File.AppendAllText(filename, exception.Message, Encoding.UTF8);
            }
            catch{ }
        }

    }
    class Program
    {
        static void Main(string[] args)
        {
            string connectionString = Properties.Settings.Default.ConnectionString;
            string table =  Properties.Settings.Default.TableName;
            string log = "";
            try 
            {
                using (SqlConnection connection = new SqlConnection(connectionString))
                {
                    connection.Open();
                    using (SqlCommand command = new SqlCommand("SELECT * FROM " + table, connection))
                    {
                        SqlDataReader reader = command.ExecuteReader();//
                        while (reader.Read())
                        {
                            DateTime timestamp = reader.GetDateTime(2);
                            string action = reader.GetString(0);  
                            string tablename = reader.GetString(1); 
                            log += timestamp.ToString() + ";Table: " + tablename + ";Action: " + action + "\r\n";
                        }
                        reader.Close();
                    }
                    SqlCommand deletion = new SqlCommand("DELETE FROM " + table, connection);
                    SqlDataReader deleter = deletion.ExecuteReader();
                    deleter.Close();
                    connection.Close();
                }
            }
            catch(Exception exception) 
            {
                Logger.Log(exception);
            }
            Logger.Memorize(log);
        }
    }

}
    