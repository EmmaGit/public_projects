using System;
using System.Linq;
using System.Management;
using System.Runtime.InteropServices;
using System.Web.Mvc;
using DWORD = System.UInt32;
using Excel = Microsoft.Office.Interop.Excel;
using LPWSTR = System.String;
using NET_API_STATUS = System.UInt32;
using System.Threading;
using System.Globalization;
using System.Web;
using System.Resources;

namespace McoEasyTool.Controllers
{
    [AllowAnonymous]
    public class HomeController : Controller
    {
        public const string AD_MODULE = "AD";
        public const string BESR_MODULE = "BESR";
        public const string APP_MODULE = "APP";
        public const string SPACE_MODULE = "SPACE";
        public static bool LOADED_SETTINGS = false;
        public const string SYSTEM_IDENTITY = "SYSTEM";
        public const string PASSWORD_HASH = "P@@Sw0rd";
        public const string SALT_KEY = "S@LT&KEY";
        public const string VI_KEY = "@1B2c3D4e5F6g7H8";


        public const string OBJECT_ATTR_ID = "Id";
        public const string OBJECT_ATTR_NAME = "Name";
        public const string OBJECT_ATTR_STATE = "State";
        public const string OBJECT_ATTR_PING = "Ping";
        public const string OBJECT_ATTR_DETAILS = "Details";
        public const string OBJECT_ATTR_DATETIME = "DateTime";
        public const string OBJECT_ATTR_DURATION = "Duration";
        public const string OBJECT_ATTR_TOTAL_ERRORS = "TotalErrors";
        public const string OBJECT_ATTR_TOTAL_CHECKED = "TotalChecked";
        public const string OBJECT_ATTR_RESULT_PATH = "ResultPath";
        public const string OBJECT_ATTR_AUTHOR = "Author";
        public const string OBJECT_ATTR_MODULE = "Module";
        public const string OBJECT_ATTR_URL = "Url";

        public const string DEFAULT_OCTECT_UNIT = "o";
        public const string DEFAULT_KILO_OCTECT_UNIT = "Ko";
        public const string DEFAULT_MEGA_OCTECT_UNIT = "Mo";
        public const string DEFAULT_GIGA_OCTECT_UNIT = "Go";
        public const string DEFAULT_TERA_OCTECT_UNIT = "To";



        //----------------------------------------------------------------------------------------
        //GENERAL FOLDERS & VARIABLES
        public static string HOSTNAME = McoToolConfig.Settings.HOSTNAME;
        public static string BATCHES_FOLDER = McoToolConfig.Settings.BATCHES_FOLDER;
        public static string LOGS_FOLDER = McoToolConfig.Settings.LOGS_FOLDER;
        public static string GENERAL_LOG_FILE = McoToolConfig.Settings.GENERAL_LOG_FILE;

        public static string DEFAULT_USERNAME_IMPERSONNATION = McoToolConfig.Settings.DEFAULT_USERNAME_IMPERSONNATION;
        public static string DEFAULT_DOMAIN_IMPERSONNATION = McoToolConfig.Settings.DEFAULT_DOMAIN_IMPERSONNATION;
        public static string DEFAULT_PASSWORD_IMPERSONNATION = McoToolConfig.Settings.DEFAULT_PASSWORD_IMPERSONNATION;
        public static string DEFAULT_DOMAIN_AND_USERNAME_IMPERSONNATION = DEFAULT_DOMAIN_IMPERSONNATION + "\\" +
            DEFAULT_USERNAME_IMPERSONNATION;

        public static string MCO_SCHEDULER_TOOL = McoToolConfig.Settings.MCO_SCHEDULER_TOOL;
        public static string MCO_IMPERSONATOR_TOOL = McoToolConfig.Settings.MCO_IMPERSONATOR_TOOL;

        public static string MAIL_ADDRESS_SENDER = McoToolConfig.Settings.MAIL_ADDRESS_SENDER;
        public static string MAIL_HOST_SERVER = McoToolConfig.Settings.MAIL_HOST_SERVER;
        public static int MAIL_PORT_NUMBER = McoToolConfig.Settings.MAIL_PORT_NUMBER;
        //END GENERAL FOLDERS & VARIABLES
        //----------------------------------------------------------------------------------------

        //AD FOLDERS & VARIABLES
        public static string AD_LOG_FILE = LOGS_FOLDER + McoToolConfig.Settings.AD_LOG_FILE;
        public static string AD_RESULTS_FOLDER = McoToolConfig.Settings.AD_RESULTS_FOLDER;
        public static string AUTO_AD_LOG_FILE = LOGS_FOLDER + McoToolConfig.Settings.AUTO_AD_LOG_FILE;

        public static int AD_MAX_REPORT_NUMBER = McoToolConfig.Settings.AD_MAX_REPORT_NUMBER; //Twice a day, about 3 Weeks


        //END AD FOLDERS & VARIABLES

        //----------------------------------------------------------------------------------------
        //BESR FOLDERS & VARIABLES
        public static int BESR_MAX_REPORT_NUMBER = McoToolConfig.Settings.BESR_MAX_REPORT_NUMBER; //Once, about 3 weeks

        public static string BESR_LOG_FILE = LOGS_FOLDER + McoToolConfig.Settings.BESR_LOG_FILE;
        public static string BESR_RESULTS_FOLDER = McoToolConfig.Settings.BESR_RESULTS_FOLDER;
        public static string AUTO_BESR_LOG_FILE = LOGS_FOLDER + McoToolConfig.Settings.AUTO_BESR_LOG_FILE;

        public static string BACKUP_REMOTE_CHECK_SERVER_ROOT_FOLDER = McoToolConfig.Settings.BACKUP_REMOTE_CHECK_SERVER_ROOT_FOLDER;
        public static string BACKUP_REMOTE_SERVER_EXEC_ROOT_FOLDER = McoToolConfig.Settings.BACKUP_REMOTE_SERVER_EXEC_ROOT_FOLDER;
        public static string BACKUP_REMOTE_SERVER_EXEC = McoToolConfig.Settings.BACKUP_REMOTE_SERVER_EXEC;

        public static string BESR_DEFAULT_INIT_FILE = BATCHES_FOLDER + McoToolConfig.Settings.BESR_DEFAULT_INIT_FILE;
        public static string BESR_RELATIVE_INIT_FILE = McoToolConfig.Settings.BESR_RELATIVE_INIT_FILE;
        public static string BESR_INIT_FILE = BATCHES_FOLDER + BESR_RELATIVE_INIT_FILE;
        public static string BESR_AUTO_UPDATE_LOG_FILE = LOGS_FOLDER + McoToolConfig.Settings.BESR_AUTO_UPDATE_LOG_FILE;
        public static string BESR_DEFAULT_INIT_FILE_README = BATCHES_FOLDER + McoToolConfig.Settings.BESR_DEFAULT_INIT_FILE_README;
        //END BESR FOLDERS & VARIABLES
        //----------------------------------------------------------------------------------------

        //----------------------------------------------------------------------------------------
        //APP FOLDERS & VARIABLES
        public static int APP_MAX_REPORT_NUMBER = McoToolConfig.Settings.APP_MAX_REPORT_NUMBER; //Once, about 3 weeks
        public static string AUTO_APP_LOG_FILE = LOGS_FOLDER + McoToolConfig.Settings.AUTO_APP_LOG_FILE;
        public static string APP_LOG_FILE = LOGS_FOLDER + McoToolConfig.Settings.APP_LOG_FILE;
        public static string APP_RESULTS_FOLDER = McoToolConfig.Settings.APP_RESULTS_FOLDER;
        public static string APP_DEFAULT_INIT_FILE = BATCHES_FOLDER + McoToolConfig.Settings.APP_DEFAULT_INIT_FILE;
        public static string APP_INIT_FILE = BATCHES_FOLDER + McoToolConfig.Settings.APP_RELATIVE_INIT_FILE;
        public static string APP_RELATIVE_INIT_FILE = McoToolConfig.Settings.APP_RELATIVE_INIT_FILE;
        public static string APP_DEFAULT_INIT_FILE_README = BATCHES_FOLDER + McoToolConfig.Settings.APP_DEFAULT_INIT_FILE_README;

        public static string[] APP_NAVIGATORS_LIST = { "IE", "FIREFOX" };
        public static string[] APP_PROCEDURE_TYPES = { "BATCH", "PROCESS", "SERVICE", "URL" };
        public static string[] APP_PROCEDURE_ACTIONS = { "START", "RESTART", "STOP", "CHECK" };
        public static string[] APP_TEXT_ATTR_TAGS_LIST = { "A", "BODY", "DIV", "LABEL", "P", "SPAN", "STRONG", "TD", "TEXTAREA", "TH" };
        public static string[] APP_SRC_ATTR_TAGS_LIST = { "AUDIO", "IMG", "SOURCE", "VIDEO" };
        public static string[] APP_VALUE_ATTR_TAGS_LIST = { "BUTTION", "INPUT", "SELECT" };
        public static string[] APP_TAGS_LIST = (APP_TEXT_ATTR_TAGS_LIST.Concat(APP_SRC_ATTR_TAGS_LIST)).Concat(APP_VALUE_ATTR_TAGS_LIST).ToArray();

        //END APP FOLDERS & VARIABLES

        //SPACE FOLDERS & VARIABLES
        public static int SPACE_MAX_REPORT_NUMBER = McoToolConfig.Settings.SPACE_MAX_REPORT_NUMBER; //Once, about 3 weeks
        public static string AUTO_SPACE_LOG_FILE = LOGS_FOLDER + McoToolConfig.Settings.AUTO_SPACE_LOG_FILE;
        public static string SPACE_LOG_FILE = LOGS_FOLDER + McoToolConfig.Settings.SPACE_LOG_FILE;
        public static string SPACE_RESULTS_FOLDER = McoToolConfig.Settings.SPACE_RESULTS_FOLDER;
        public static string SPACE_DEFAULT_INIT_FILE = BATCHES_FOLDER + McoToolConfig.Settings.SPACE_DEFAULT_INIT_FILE;
        public static string SPACE_INIT_FILE = BATCHES_FOLDER + McoToolConfig.Settings.SPACE_RELATIVE_INIT_FILE;
        public static string SPACE_RELATIVE_INIT_FILE = McoToolConfig.Settings.SPACE_RELATIVE_INIT_FILE;
        public static string SPACE_DEFAULT_INIT_FILE_README = BATCHES_FOLDER + McoToolConfig.Settings.SPACE_DEFAULT_INIT_FILE_README;
        public static double SPACE_DEFAULT_THRESHOLD = McoToolConfig.Settings.SPACE_DEFAULT_THRESHOLD;
        public static string DEFAULT_NUMBER_DECIMAL_SEPARATOR = McoToolConfig.Settings.DEFAULT_NUMBER_DECIMAL_SEPARATOR;
        public static int SPACE_MAX_CHARTS_LINES_NUMBER = McoToolConfig.Settings.SPACE_MAX_CHARTS_LINES_NUMBER;
        public static double DEFAULT_OCTECT_INCREMENT = McoToolConfig.Settings.DEFAULT_OCTECT_INCREMENT;

        public static string SPACE_USERNAME_IMPERSONNATION = McoToolConfig.Settings.SPACE_USERNAME_IMPERSONNATION;
        public static string SPACE_DOMAIN_IMPERSONNATION = McoToolConfig.Settings.SPACE_DOMAIN_IMPERSONNATION;
        public static string SPACE_PASSWORD_IMPERSONNATION = McoToolConfig.Settings.SPACE_PASSWORD_IMPERSONNATION;
        public static string SPACE_DOMAIN_AND_USERNAME_IMPERSONNATION = SPACE_DOMAIN_IMPERSONNATION + "\\" +
            SPACE_USERNAME_IMPERSONNATION;

        public static string SPACE_DEFAULT_LOCAL_MAPPED_DRIVE_LETTER = BATCHES_FOLDER + McoToolConfig.Settings.SPACE_DEFAULT_LOCAL_MAPPED_DRIVE_LETTER;
        public static double LOG_SIZE_LIMIT = McoToolConfig.Settings.LOG_SIZE_LIMIT;

        //END SPACE FOLDERS & VARIABLES
        //----------------------------------------------------------------------------------------

        /*[HttpGet]
        public ActionResult Index(string cultureName = null)
        {
            //Modify current thread's culture  
            ViewBag.Message = "Bienvenue dans l'application d'automatisation des procédures";
            return View();
        }*/

        public ActionResult Index()
        {
            ViewBag.Message = "Bienvenue dans l'application d'automatisation des procédures";

            return View();
        }

        public ActionResult Home()
        {
            ViewBag.Message = "Bienvenue dans l'application d'automatisation des procédures";

            return View();
        }

        public ActionResult Contact()
        {
            ViewBag.Message = "Contacts.";

            return View();
        }

        public ActionResult Utilities()
        {
            //ViewBag.Message = Decrypt(SPACE_PASSWORD_IMPERSONNATION); //"Boîte à outils";
            ViewBag.Message = "Boîte à outils";
            return View();
        }

        [HttpPost]
        public string Encrypt()
        {
            string password = "";
            try
            {
                password = Request.Form["password"].ToString();
                password = McoUtilities.Encrypt(password);
                return password;
            }
            catch (Exception exception)
            {
                McoUtilities.General_Logging(exception, "Encrypt", 0, User.Identity.Name);

            }
            return "<Ne pas copier> Une erreur est survenue </Ne pas copier>";
        }

        private string Decrypt(string encryptedText)
        {
            return McoUtilities.Decrypt(encryptedText);
        }

    }
}
