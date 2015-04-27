using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.SqlClient;
using System.Configuration;
using System.Net.Mail;

namespace TRG_DBA_UPDATE
{
    public class Logger
    {
        public static void Memorize(string content) 
        {
            string filename = Properties.Settings.Default.MEMO_FILE;
            try 
            {
                System.IO.File.AppendAllText(filename, content, Encoding.UTF8);
            }
            catch(Exception exception)
            {
                Log(exception);
            }
        }

        public static void Log(Exception exception = null) 
        {
            string filename = Properties.Settings.Default.LOG_FILE;
            string extension = filename.Substring(filename.LastIndexOf("."));
            string content = DateTime.Now.ToString();
            string separator = ";";
            if(exception == null)
            {
                exception = new Exception("No message");
            }
            try
            {
                switch(extension)
                {
                    case ".csv": separator = ";"; break;
                    case ".txt": separator = "\t\t"; break;
                    case ".log": separator = "\t\t"; break;
                    default: separator = ";"; break;
                }
                content += separator + "\"" + exception.Message + separator + "\"" + exception.Source + "\r\n";
                System.IO.File.AppendAllText(filename, exception.Message, Encoding.UTF8);
            }
            catch{ }
        }

    }
    class Program
    {
        static void Main(string[] args)
        {
            string connectionString = Properties.Settings.Default.CONNECTION_STRING;
            string table =  Properties.Settings.Default.TABLE_NAME;
            string log = "";
            try 
            {
                using (SqlConnection connection = new SqlConnection(connectionString))
                {
                    Console.WriteLine("Database connection attempt...");
                    connection.Open();
                    Console.WriteLine("Database connection established");
                    using (SqlCommand command = new SqlCommand("SELECT * FROM " + table, connection))
                    {
                        Console.WriteLine("SQL quering execution....");
                        SqlDataReader reader = command.ExecuteReader();//
                        Console.WriteLine("Reading of the results...");
                        while (reader.Read())
                        {
                            DateTime timestamp = reader.GetDateTime(2);
                            string action = reader.GetString(0);  
                            string tablename = reader.GetString(1); 
                            log += timestamp.ToString() + ";Table: " + tablename + ";Action: " + action + "\r\n";
                        }
                        reader.Close();
                        Console.WriteLine("Reading achieved");
                    }
                    if(log.Trim() == "")
                    {
                        Console.WriteLine("Empty Table, alerting...");
                        Alert();
                    }
                    else
                    {
                        Console.WriteLine("Droping table for next occurence...");
                        SqlCommand deletion = new SqlCommand("DELETE FROM " + table, connection);
                        SqlDataReader deleter = deletion.ExecuteReader();
                        deleter.Close();
                    }
                    connection.Close();
                    Console.WriteLine("SQL connection closed");
                }
            }
            catch(Exception exception) 
            {
                Console.Error.WriteLine("Damn!! \r\n" + exception.Message);
                Logger.Log(exception);
            }
            Logger.Memorize(log);
        }

        static void Alert() 
        {
            try 
            {
                SmtpClient client = new SmtpClient();
                client.Port = Properties.Settings.Default.MAIL_PORT_NUMBER;
                client.DeliveryMethod = SmtpDeliveryMethod.Network;
                client.UseDefaultCredentials = false;
                client.Host = Properties.Settings.Default.MAIL_HOST_SERVER;

                string body = "No entry were found on " + DateTime.Now.ToString() + "<br/>" +
                    "Please, check that everything is Okay";

                body = DateTime.Now.ToString() + ":<br/>" + Properties.Settings.Default.MAIL_BODY;

                MailMessage mail = new MailMessage();
                mail.From = new MailAddress(Properties.Settings.Default.MAIL_ADDRESS_SENDER);
                mail.IsBodyHtml = true;
                mail.Body = body;
                string[] recipients = Properties.Settings.Default.MAIL_RECIPIENTS_LIST.Split(';');
                foreach (string recipient in recipients)
                {
                    if (recipient.Trim() == "")
                    {
                        continue;
                    }
                    mail.To.Add(new MailAddress(recipient.Trim()));
                }
                //mail.Attachments.Add(new Attachment(@email.Report.ResultPath));
                mail.Subject = "History table update " + DateTime.Now.ToString();
                // Send.
                client.Send(mail);
                Console.WriteLine("Email sent!!!");
            }
            catch (Exception exception) 
            {
                Console.Error.WriteLine("Damn!! \r\n" + exception.Message);
                Logger.Log(exception);
            }  
        }
    }

}
    