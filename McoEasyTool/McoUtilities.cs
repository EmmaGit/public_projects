using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Linq.Expressions;
using System.Management;
using System.Net.Mail;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.Security.Principal;
using System.Text;
using LPWSTR = System.String;
using System.Web.Mvc;
using DWORD = System.UInt32;
using Excel = Microsoft.Office.Interop.Excel;
using NET_API_STATUS = System.UInt32;


namespace McoEasyTool.Controllers
{
    public class McoUtilitiesController : Controller
    {
        private static DataModelContainer db = new DataModelContainer();

        public Email SetRecipients(Email email)
        {
            ICollection<Recipient> recipients = db.Recipients.Where(rec => rec.Module == email.Module)
                .Where(inc => inc.Included != false).ToList();
            try
            {
                foreach (Recipient recipient in recipients)
                {
                    if (recipient.AbsoluteAddress != null && recipient.AbsoluteAddress.Trim() != "")
                    {
                        email.Recipients += recipient.AbsoluteAddress + "; ";
                    }
                    else
                    {
                        email.Recipients += recipient.RelativeAddress + "@bouygues-construction.com; ";
                    }
                }
            }
            catch { }
            return email;
        }

        public JsonResult InternalSend(Email email)
        {
            try
            {
                Dictionary<string, string> result = new Dictionary<string, string>();
                result.Add("Response", "OK");
                result.Add("Email", email.Id.ToString());
                result.Add("Report", email.Report.Id.ToString());
                result.Add("applications", "");
                result.Add("errors", "");
                result.Add("email", email.Id.ToString());
                result.Add("response", "OK");
                result.Add("content", "");

                SmtpClient client = new SmtpClient();
                client.Port = HomeController.MAIL_PORT_NUMBER;
                client.DeliveryMethod = SmtpDeliveryMethod.Network;
                client.UseDefaultCredentials = false;
                client.Host = HomeController.MAIL_HOST_SERVER;


                MailMessage mail = new MailMessage();
                mail.From = new MailAddress(HomeController.MAIL_ADDRESS_SENDER);
                mail.IsBodyHtml = true;
                mail.Body = email.Body;
                mail.Subject = email.Subject;
                string[] recipientsList = email.Recipients.Split(';');
                foreach (string recipient in recipientsList)
                {
                    if (recipient.Trim() == "")
                    {
                        continue;
                    }
                    mail.To.Add(new MailAddress(recipient.Trim()));
                }
                // Send.
                client.Send(mail);
                email.Sent = true;
                result["Response"] = "Emergency email sent.";
                McoUtilities.General_Logging(new Exception("...."), "InternalSend Email", 3);
                return Json(result, JsonRequestBehavior.AllowGet);
            }
            catch (Exception exception)
            {
                McoUtilities.General_Logging(exception, "InternalSend Email", 0);
                Dictionary<string, string> result = new Dictionary<string, string>();
                result.Add("Response", "Emergency email unsent");
                result.Add("Error", exception.Message);
                return Json(result, JsonRequestBehavior.AllowGet);
            }
        }

        public JsonResult InternalOpen(Email email)
        {
            Dictionary<string, string> result = new Dictionary<string, string>();
            result.Add("Response", "OK");
            result.Add("Email", email.Id.ToString());
            result.Add("Report", email.Report.Id.ToString());
            result.Add("applications", "N/A");
            result.Add("errors", "Can't check");
            result.Add("email", email.Id.ToString());
            result.Add("response", "OK");
            result.Add("content", "");
            result.Add("subject", email.Subject);
            result.Add("recipient", email.Recipients);
            result.Add("body", email.Body);
            result.Add("servers", "N/A");

            return Json(result, JsonRequestBehavior.AllowGet);
        }
    }

    public static class McoUtilities
    {

        public static UNCAccessWithCredentials UNC_ACCESS = new UNCAccessWithCredentials();

        private static DataModelContainer db = new DataModelContainer();

        public static IEnumerable<TEntity> OrderBy<TEntity>(this IEnumerable<TEntity> source,
                                                    string orderByProperty, bool desc)
        {
            string command = desc ? "OrderByDescending" : "OrderBy";
            var type = typeof(TEntity);
            var property = type.GetProperty(orderByProperty);
            var parameter = Expression.Parameter(type, "p");
            var propertyAccess = Expression.MakeMemberAccess(parameter, property);
            var orderByExpression = Expression.Lambda(propertyAccess, parameter);
            var resultExpression = Expression.Call(typeof(Queryable), command,
                                                   new[] { type, property.PropertyType },
                                                   source.AsQueryable().Expression,
                                                   Expression.Quote(orderByExpression));
            return source.AsQueryable().Provider.CreateQuery<TEntity>(resultExpression);
        }

        public static T[] GetBoundaries<T>(T item, string orderByProperty, bool desc)
        {
            if (item == null)
            {
                return null;
            }
            T[] Boundaries = new T[5];
            List<T> items = new List<T>();
            int index = 0;

            Boundaries[2] = item;
            try
            {
                items = db.Set(item.GetType()).AsQueryable().OfType<T>().ToList();
            }
            catch { }
            try
            {
                items = OrderBy<T>(items, orderByProperty, desc).ToList();
            }
            catch { }
            object item_id = GetPropertyValue(item, HomeController.OBJECT_ATTR_ID);
            try
            {
                foreach (T obj in items)
                {
                    object obj_id = GetPropertyValue(obj, HomeController.OBJECT_ATTR_ID);
                    if (obj_id.Equals(item_id))
                    {
                        break;
                    }
                    index++;
                }
            }
            catch { }
            if (index - 1 >= 0)
            {
                Boundaries[1] = items.ElementAt(index - 1);
            }
            if (index + 1 <= items.Count - 1)
            {
                Boundaries[3] = items.ElementAt(index + 1);
            }
            if (Boundaries[1] != null && index - 1 != 0)
            {
                Boundaries[0] = items.ElementAt(0);
            }
            if (Boundaries[3] != null && index + 1 != items.Count - 1)
            {
                Boundaries[4] = items.ElementAt(items.Count - 1);
            }
            return Boundaries;
        }

        public static object[] GetBoundariesInfo<T>(T item, string returnedProperty, string orderByProperty, bool desc = false)
        {
            T[] boundaries = GetBoundaries<T>(item, orderByProperty, desc);
            object[] values = new object[5];
            var type = typeof(T);
            PropertyInfo property = null;
            try
            {
                property = type.GetProperty(returnedProperty);
                values[0] = (boundaries[0] != null) ? property.GetValue(boundaries[0], null) : null;
                values[1] = (boundaries[1] != null) ? property.GetValue(boundaries[1], null) : null;
                values[2] = (boundaries[2] != null) ? property.GetValue(boundaries[2], null) : null;
                values[3] = (boundaries[3] != null) ? property.GetValue(boundaries[3], null) : null;
                values[4] = (boundaries[4] != null) ? property.GetValue(boundaries[4], null) : null;
            }
            catch { }
            return values;
        }

        public static object GetPropertyValue<T>(T item, string returnedProperty)
        {
            object value = null;
            var type = typeof(T);
            PropertyInfo property = null;
            try
            {
                property = type.GetProperty(HomeController.OBJECT_ATTR_ID);
                value = property.GetValue(item, null);
            }
            catch { }
            return value;
        }

        public static object[] GetIdValues<T>(T item, string orderByProperty, bool desc = false)
        {
            object[] boundaries =
                    GetBoundariesInfo<T>(item,
                        HomeController.OBJECT_ATTR_ID, orderByProperty, desc);
            int index = 0;
            foreach (object obj in boundaries)
            {
                if (obj == null)
                {
                    boundaries[index] = 0;
                }
                index++;
            }
            return boundaries;
        }

        public static string ChangeSPChart(string sTheInput)
        {
            StringBuilder sRetMe = new StringBuilder(sTheInput);

            sRetMe.Replace('+', '-');
            sRetMe.Replace('/', '*');
            sRetMe.Replace('=', '!');

            return sRetMe.ToString();
        }

        public static string FixSPChart(string sTheInput)
        {
            StringBuilder sRetMe = new StringBuilder(sTheInput);

            sRetMe.Replace('-', '+');
            sRetMe.Replace('*', '/');
            sRetMe.Replace('!', '=');

            return sRetMe.ToString();
        }

        public static string Encrypt(string plainText)
        {
            byte[] plainTextBytes = Encoding.UTF8.GetBytes(plainText);

            byte[] keyBytes = new Rfc2898DeriveBytes(HomeController.PASSWORD_HASH,
                Encoding.ASCII.GetBytes(HomeController.SALT_KEY)).GetBytes(256 / 8);
            var symmetricKey = new RijndaelManaged() { Mode = CipherMode.CBC, Padding = PaddingMode.Zeros };
            var encryptor = symmetricKey.CreateEncryptor(keyBytes,
                Encoding.ASCII.GetBytes(HomeController.VI_KEY));

            byte[] cipherTextBytes;

            using (var memoryStream = new MemoryStream())
            {
                using (var cryptoStream = new CryptoStream(memoryStream, encryptor, CryptoStreamMode.Write))
                {
                    cryptoStream.Write(plainTextBytes, 0, plainTextBytes.Length);
                    cryptoStream.FlushFinalBlock();
                    cipherTextBytes = memoryStream.ToArray();
                    cryptoStream.Close();
                }
                memoryStream.Close();
            }
            return ChangeSPChart(Convert.ToBase64String(cipherTextBytes));
        }

        public static string Decrypt(string encryptedText)
        {
            encryptedText = FixSPChart(encryptedText);
            byte[] cipherTextBytes = Convert.FromBase64String(encryptedText);
            byte[] keyBytes = new Rfc2898DeriveBytes(HomeController.PASSWORD_HASH,
                Encoding.ASCII.GetBytes(HomeController.SALT_KEY)).GetBytes(256 / 8);
            var symmetricKey = new RijndaelManaged() { Mode = CipherMode.CBC, Padding = PaddingMode.None };

            var decryptor = symmetricKey.CreateDecryptor(keyBytes, Encoding.ASCII.GetBytes(HomeController.VI_KEY));
            var memoryStream = new MemoryStream(cipherTextBytes);
            var cryptoStream = new CryptoStream(memoryStream, decryptor, CryptoStreamMode.Read);
            byte[] plainTextBytes = new byte[cipherTextBytes.Length];

            int decryptedByteCount = cryptoStream.Read(plainTextBytes, 0, plainTextBytes.Length);
            memoryStream.Close();
            cryptoStream.Close();
            return Encoding.UTF8.GetString(plainTextBytes, 0, decryptedByteCount).TrimEnd("\0".ToCharArray());
        }

        [DllImport("advapi32.dll", SetLastError = true)]
        public static extern bool LogonUser(
                string lpszUsername,
                string lpszDomain,
                string lpszPassword,
                int dwLogonType,
                int dwLogonProvider,
                out IntPtr phToken);
        public enum LogonType
        {
            /// <summary>
            /// This logon type is intended for users who will be interactively using the computer, such as a user being logged on  
            /// by a terminal server, remote shell, or similar process.
            /// This logon type has the additional expense of caching logon information for disconnected operations; 
            /// therefore, it is inappropriate for some client/server applications,
            /// such as a mail server.
            /// </summary>
            LOGON32_LOGON_INTERACTIVE = 2,

            /// <summary>
            /// This logon type is intended for high performance servers to authenticate plaintext passwords.

            /// The LogonUser function does not cache credentials for this logon type.
            /// </summary>
            LOGON32_LOGON_NETWORK = 3,

            /// <summary>
            /// This logon type is intended for batch servers, where processes may be executing on behalf of a user without 
            /// their direct intervention. This type is also for higher performance servers that process many plaintext
            /// authentication attempts at a time, such as mail or Web servers. 
            /// The LogonUser function does not cache credentials for this logon type.
            /// </summary>
            LOGON32_LOGON_BATCH = 4,

            /// <summary>
            /// Indicates a service-type logon. The account provided must have the service privilege enabled. 
            /// </summary>
            LOGON32_LOGON_SERVICE = 5,

            /// <summary>
            /// This logon type is for GINA DLLs that log on users who will be interactively using the computer. 
            /// This logon type can generate a unique audit record that shows when the workstation was unlocked. 
            /// </summary>
            LOGON32_LOGON_UNLOCK = 7,

            /// <summary>
            /// This logon type preserves the name and password in the authentication package, which allows the server to make 
            /// connections to other network servers while impersonating the client. A server can accept plaintext credentials 
            /// from a client, call LogonUser, verify that the user can access the system across the network, and still 
            /// communicate with other servers.
            /// NOTE: Windows NT:  This value is not supported. 
            /// </summary>
            LOGON32_LOGON_NETWORK_CLEARTEXT = 8,

            /// <summary>
            /// This logon type allows the caller to clone its current token and specify new credentials for outbound connections.
            /// The new logon session has the same local identifier but uses different credentials for other network connections. 
            /// NOTE: This logon type is supported only by the LOGON32_PROVIDER_WINNT50 logon provider.
            /// NOTE: Windows NT:  This value is not supported. 
            /// </summary>
            LOGON32_LOGON_NEW_CREDENTIALS = 9,
        }

        public enum LogonProvider
        {
            /// <summary>
            /// Use the standard logon provider for the system. 
            /// The default security provider is negotiate, unless you pass NULL for the domain name and the user name 
            /// is not in UPN format. In this case, the default provider is NTLM. 
            /// NOTE: Windows 2000/NT:   The default security provider is NTLM.
            /// </summary>
            LOGON32_PROVIDER_DEFAULT = 0,
            LOGON32_PROVIDER_WINNT35 = 1,
            LOGON32_PROVIDER_WINNT40 = 2,
            LOGON32_PROVIDER_WINNT50 = 3
        }

        public class UNCAccessWithCredentials : IDisposable
        {
            [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
            internal struct USE_INFO_2
            {
                internal LPWSTR ui2_local;
                internal LPWSTR ui2_remote;
                internal LPWSTR ui2_password;
                internal DWORD ui2_status;
                internal DWORD ui2_asg_type;
                internal DWORD ui2_refcount;
                internal DWORD ui2_usecount;
                internal LPWSTR ui2_username;
                internal LPWSTR ui2_domainname;
            }

            [DllImport("NetApi32.dll", SetLastError = true, CharSet = CharSet.Unicode)]
            internal static extern NET_API_STATUS NetUseAdd(
                LPWSTR UncServerName,
                DWORD Level,
                ref USE_INFO_2 Buf,
                out DWORD ParmError);

            [DllImport("NetApi32.dll", SetLastError = true, CharSet = CharSet.Unicode)]
            internal static extern NET_API_STATUS NetUseDel(
                LPWSTR UncServerName,
                LPWSTR UseName,
                DWORD ForceCond);

            private bool disposed = false;

            private string sUNCPath;
            private string sUser;
            private string sPassword;
            private string sDomain;
            private int iLastError;

            /// <summary>
            /// A disposeable class that allows access to a UNC resource with credentials.
            /// </summary>
            public UNCAccessWithCredentials()
            {
            }

            /// <summary>
            /// The last system error code returned from NetUseAdd or NetUseDel.  Success = 0
            /// </summary>
            public int LastError
            {
                get { return iLastError; }
            }

            public void Dispose()
            {
                if (!this.disposed)
                {
                    NetUseDelete();
                }
                disposed = true;
                GC.SuppressFinalize(this);
            }

            /// <summary>
            /// Connects to a UNC path using the credentials supplied.
            /// </summary>
            /// <param name="UNCPath">Fully qualified domain name UNC path</param>
            /// <param name="User">A user with sufficient rights to access the path.</param>
            /// <param name="Domain">Domain of User.</param>
            /// <param name="Password">Password of User</param>
            /// <returns>True if mapping succeeds.  Use LastError to get the system error code.</returns>
            public bool NetUseWithCredentials(string UNCPath, string User, string Domain, string Password)
            {
                sUNCPath = UNCPath;
                sUser = User;
                sPassword = Password;
                sDomain = Domain;
                return NetUseWithCredentials();
            }

            private bool NetUseWithCredentials()
            {
                uint returncode;
                try
                {
                    USE_INFO_2 useinfo = new USE_INFO_2();

                    useinfo.ui2_remote = sUNCPath;
                    useinfo.ui2_username = sUser;
                    useinfo.ui2_domainname = sDomain;
                    useinfo.ui2_password = sPassword;
                    useinfo.ui2_asg_type = 0;
                    useinfo.ui2_usecount = 1;
                    uint paramErrorIndex;
                    returncode = NetUseAdd(null, 2, ref useinfo, out paramErrorIndex);
                    iLastError = (int)returncode;
                    return returncode == 0;
                }
                catch
                {
                    iLastError = Marshal.GetLastWin32Error();
                    return false;
                }
            }

            /// <summary>
            /// Ends the connection to the remote resource 
            /// </summary>
            /// <returns>True if it succeeds.  Use LastError to get the system error code</returns>
            public bool NetUseDelete()
            {
                uint returncode;
                try
                {
                    returncode = NetUseDel(null, sUNCPath, 2);
                    iLastError = (int)returncode;
                    return (returncode == 0);
                }
                catch
                {
                    iLastError = Marshal.GetLastWin32Error();
                    return false;
                }
            }

            ~UNCAccessWithCredentials()
            {
                Dispose();
            }
        }

        public static bool CloseExcel(Excel.Application MyApplication, Excel.Workbook MyWorkbook, Excel.Worksheet MySheet)
        {
            if (MySheet != null)
            {
                Marshal.FinalReleaseComObject(MySheet);
                MySheet = null;
            }

            if (MyWorkbook != null)
            {
                MyWorkbook.Close();
                Marshal.FinalReleaseComObject(MyWorkbook);
                MyWorkbook = null;
            }

            if (MyApplication != null)
            {
                Marshal.FinalReleaseComObject(MyApplication.Workbooks);
                MyApplication.Quit();
                Marshal.FinalReleaseComObject(MyApplication);
                MyApplication = null;
            }

            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect();
            GC.WaitForPendingFinalizers();
            return true;
        }

        public static string GetProcessOwner(int processId)
        {
            string query = "Select * From Win32_Process Where ProcessID = " + processId;
            ManagementObjectSearcher moSearcher = new ManagementObjectSearcher(query);
            ManagementObjectCollection moCollection = moSearcher.Get();

            foreach (ManagementObject mo in moCollection)
            {
                string[] args = new string[] { string.Empty };
                int returnVal = Convert.ToInt32(mo.InvokeMethod("GetOwner", args));
                if (returnVal == 0)
                    return args[0];
            }
            return "N/A";
        }

        //0 error
        //1 warning
        //2 info
        public static void General_Logging(Exception exception, string action, int level = 0, string author = "UNKNOWN")
        {
            string log = "";
            switch (level)
            {
                case 0: log = "[ERROR]"; break;
                case 2: log = "[WARNING]"; break;
                case 3: log = "[INFO]"; break;
                default: log = "[UNKNOWN]"; break;
            }
            try
            {
                log = "\"" + DateTime.Now.ToString() + "\";\"" + log + "\";\"" + action + "\";\"";
                log += (exception.InnerException == null) ? exception.Message + "\";\"" + author + "\"\r\n" :
                    exception.Message + " || " + exception.InnerException + "\";\"" + author + "\"\r\n";
                System.IO.File.AppendAllText(HomeController.LOGS_FOLDER + HomeController.GENERAL_LOG_FILE, log, Encoding.UTF8);
            }
            catch { }
        }

        public static void Specific_Logging(Exception exception, string action, string module, int level = 0, string author = "UNKNOWN")
        {
            string log = "", filename = "";
            switch (level)
            {
                case 0: log = "[ERROR]"; break;
                case 2: log = "[WARNING]"; break;
                case 3: log = "[INFO]"; break;
                default: log = "[UNKNOWN]"; break;
            }
            switch (module)
            {
                case HomeController.AD_MODULE: filename = HomeController.AD_LOG_FILE; break;
                case HomeController.BESR_MODULE: filename = HomeController.BESR_LOG_FILE; break;
                case HomeController.APP_MODULE: filename = HomeController.APP_LOG_FILE; break;
                case HomeController.SPACE_MODULE: filename = HomeController.SPACE_LOG_FILE; break;
            }
            try
            {
                log = "\"" + DateTime.Now.ToString() + "\";\"" + log + "\";\"" + action + "\";\"";
                log += (exception.InnerException == null) ? exception.Message + "\";\"" + author + "\"\r\n" :
                    exception.Message + " || " + exception.InnerException + "\";\"" + author + "\"\r\n";
                System.IO.File.AppendAllText(filename, log, Encoding.UTF8);
            }
            catch { }
            General_Logging(exception, action, level, author);
        }

        public static string GetModuleDescription(string module, int number = 0)
        {
            string filename = "", lines = "";
            switch (module)
            {
                case HomeController.AD_MODULE: filename = "Active_Directory_Desc_" +
                    number.ToString() + ".html"; break;
                case HomeController.BESR_MODULE: filename = "Backup_Exec_Desc_" +
                    number.ToString() + ".html"; break;
                case HomeController.APP_MODULE: filename = "Flash_Test_Desc_" +
                    number.ToString() + ".html"; break;
                case HomeController.SPACE_MODULE: filename = "Capacity_Planning_Desc_" +
                    number.ToString() + ".html"; break;
                default: filename = "Mco_Easy_Tool_Desc.html"; break;
            }
            string[] content = System.IO.File.ReadAllLines(HomeController.BATCHES_FOLDER + filename);
            foreach (string line in content)
            {
                lines += line + "<br/>";
            }
            return lines;
        }

        public static bool IsValidLoginPassword(string username, string password)
        {
            string[] infos = username.Split(new string[] { @"\" }, StringSplitOptions.RemoveEmptyEntries);
            string domain = infos[0];
            username = infos[1];
            IntPtr userToken = IntPtr.Zero;
            bool success = LogonUser(
                    username,
                    domain,
                    McoUtilities.Decrypt(password),
                    (int)LogonType.LOGON32_LOGON_INTERACTIVE, //2
                    (int)LogonProvider.LOGON32_PROVIDER_DEFAULT, //0
                    out userToken);
            return success;
        }

        /*public static bool CreateRdpFile(string servername)
        {
            string[] lines = System.IO.File.ReadAllLines(HomeController.DEFAULT_RDP_FILE, Encoding.Default);
            string[] content = new string[lines.Length];
            int index = 0;
            foreach (string line in lines)
            {
                if (line.IndexOf(HomeController.DEFAULT_RDP_HOSTNAME_KEY) != -1)
                {
                    content[index] = line.Replace(HomeController.DEFAULT_RDP_HOSTNAME_KEY, servername.ToUpper());
                }
                else
                {
                    content[index] = line;
                }

                index++;
            }
            try
            {
                System.IO.File.AppendAllLines(HomeController.BATCHES_FOLDER + servername + ".rdp", content, Encoding.Default);
                General_Logging(new Exception("..."), "CreateRdpFile " + servername);
                return true;
            }
            catch (Exception exception)
            {
                General_Logging(exception, "CreateRdpFile " + servername);
                return false;
            }
        }

        public static Process OpenRdpConnection(string servername)
        {
            try
            {
                Process process = new Process();
                process.StartInfo.FileName = Environment.ExpandEnvironmentVariables(@"%SystemRoot%\system32\cmdkey.exe");
                process.StartInfo.Arguments = "/generic:TERMSRV/" + servername + " /user:" +
                    HomeController.DEFAULT_DOMAIN_AND_USERNAME_IMPERSONNATION + " /pass:" +
                    Decrypt(HomeController.DEFAULT_PASSWORD_IMPERSONNATION);
                process.Start();

                process.StartInfo.FileName = Environment.ExpandEnvironmentVariables(@"%SystemRoot%\system32\mstsc.exe");
                process.StartInfo.Arguments = "/v " + servername; // ip or name of computer to connect
                process.Start();
                General_Logging(new Exception(), "OpenRdpConnection " + servername + " Process Id:" + process.Id);
                return process;
            }
            catch (Exception exception)
            {
                General_Logging(exception, "OpenRdpConnection " + servername);
                return null;
            }
        }

        public static bool CloseRdpConnection(Process process)
        {
            try
            {
                process.Close();
                return true;
            }
            catch (Exception exception)
            {
                General_Logging(exception, "CloseRdpConnection " + process.Id);
                return false;
            }
        }*/

        public static bool CheckIfLoggedOn(string username)
        {
            var ntuser = new NTAccount(username);

            var securityIdentifier = (SecurityIdentifier)ntuser.Translate(typeof(SecurityIdentifier));

            var okey = Registry.Users.OpenSubKey(securityIdentifier + @"\Control Panel\Desktop", true); //any key is ok

            if (okey != null)
                return true;
            else
                return false;
        }

        public static JsonResult NotifyImpossibility(string module, bool autosend = true)
        {
            McoUtilitiesController utilities_Controller = new McoUtilitiesController();
            Report report = db.Reports.Create();
            string subject = "", body = "";
            Email email = db.Emails.Create();
            email.Module = module;
            email.Recipients = "";
            switch (module)
            {
                case HomeController.AD_MODULE:
                    subject = "Erreur lors de la génération du rapport AD " + DateTime.Now;
                    body = "<br/><strong>Repadmin.exe</strong> est actuellement en cours d'utilisation.";
                    break;
                case HomeController.APP_MODULE:
                    subject = "Erreur lors de la génération du Flash Test : " + DateTime.Now;
                    body = "<br/>Pour que la simulation d'authentification du flash test fonctionne correctement," +
                        " il faut qu'une session soit ouverte sur le <strong>BCNVSRV192</strong> via le compte <strong>" +
                        HomeController.DEFAULT_DOMAIN_AND_USERNAME_IMPERSONNATION + "</stroong>.<br />Prière d'ouvrir cette session et de relancer le check manuellement.<br />";
                    break;
                case HomeController.BESR_MODULE: break;
                case HomeController.SPACE_MODULE: break;
            }
            report.DateTime = DateTime.Now;
            report.TotalChecked = 0;
            report.TotalErrors = 0;
            report.ResultPath = "N/A:Error";
            report.Module = module;
            report.Author = HomeController.SYSTEM_IDENTITY;

            report.Email = email;
            email.Report = report;
            email.Subject = subject;
            email.Body = body;
            email = utilities_Controller.SetRecipients(email);
            report.Duration = DateTime.Now.Subtract(report.DateTime);
            db.Reports.Add(report);
            db.SaveChanges();
            int emailId = emailId = report.Email.Id;


            if (!autosend)
            {
                return utilities_Controller.InternalOpen(email);
            }
            return utilities_Controller.InternalSend(email);
        }

    }
}