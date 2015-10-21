using System.Configuration;
using System.Globalization;
using System;
namespace McoEasyTool
{
    public class McoToolConfig : ConfigurationSection
    {
        private static readonly McoToolConfig ConfigSection = ConfigurationManager.GetSection("McoToolConfig") as McoToolConfig;

        public static McoToolConfig Settings
        {
            get
            {
                return ConfigSection;
            }
        }

        [ConfigurationProperty("HOSTNAME", IsRequired = true)]
        public string HOSTNAME
        {
            get
            {
                return this["HOSTNAME"] as string;
            }
        }

        [ConfigurationProperty("BATCHES_FOLDER", IsRequired = true)]
        public string BATCHES_FOLDER
        {
            get
            {
                return this["BATCHES_FOLDER"] as string;
            }
        }

        [ConfigurationProperty("LOGS_FOLDER", IsRequired = true)]
        public string LOGS_FOLDER
        {
            get
            {
                return this["LOGS_FOLDER"] as string;
            }
        }

        [ConfigurationProperty("GENERAL_LOG_FILE", IsRequired = true)]
        public string GENERAL_LOG_FILE
        {
            get
            {
                return this["GENERAL_LOG_FILE"] as string;
            }
        }

        [ConfigurationProperty("DEFAULT_USERNAME_IMPERSONNATION", IsRequired = true)]
        public string DEFAULT_USERNAME_IMPERSONNATION
        {
            get
            {
                return this["DEFAULT_USERNAME_IMPERSONNATION"] as string;
            }
        }

        [ConfigurationProperty("DEFAULT_DOMAIN_IMPERSONNATION", IsRequired = true)]
        public string DEFAULT_DOMAIN_IMPERSONNATION
        {
            get
            {
                return this["DEFAULT_DOMAIN_IMPERSONNATION"] as string;
            }
        }

        [ConfigurationProperty("DEFAULT_PASSWORD_IMPERSONNATION", IsRequired = true)]
        public string DEFAULT_PASSWORD_IMPERSONNATION
        {
            get
            {
                return this["DEFAULT_PASSWORD_IMPERSONNATION"] as string;
            }
        }

        [ConfigurationProperty("MCO_SCHEDULER_TOOL", IsRequired = true)]
        public string MCO_SCHEDULER_TOOL
        {
            get
            {
                return this["MCO_SCHEDULER_TOOL"] as string;
            }
        }

        [ConfigurationProperty("MCO_IMPERSONATOR_TOOL", IsRequired = true)]
        public string MCO_IMPERSONATOR_TOOL
        {
            get
            {
                return this["MCO_IMPERSONATOR_TOOL"] as string;
            }
        }

        [ConfigurationProperty("MAIL_ADDRESS_SENDER", IsRequired = true)]
        public string MAIL_ADDRESS_SENDER
        {
            get
            {
                return this["MAIL_ADDRESS_SENDER"] as string;
            }
        }

        [ConfigurationProperty("MAIL_HOST_SERVER", IsRequired = true)]
        public string MAIL_HOST_SERVER
        {
            get
            {
                return this["MAIL_HOST_SERVER"] as string;
            }
        }

        [ConfigurationProperty("MAIL_PORT_NUMBER", IsRequired = true)]
        public int MAIL_PORT_NUMBER
        {
            get
            {
                int port = 25;
                int.TryParse(this["MAIL_PORT_NUMBER"].ToString(), out port);
                return port;
            }
        }

        [ConfigurationProperty("AD_LOG_FILE", IsRequired = true)]
        public string AD_LOG_FILE
        {
            get
            {
                return this["AD_LOG_FILE"] as string;
            }
        }

        [ConfigurationProperty("AD_RESULTS_FOLDER", IsRequired = true)]
        public string AD_RESULTS_FOLDER
        {
            get
            {
                return this["AD_RESULTS_FOLDER"] as string;
            }
        }

        [ConfigurationProperty("AUTO_AD_LOG_FILE", IsRequired = true)]
        public string AUTO_AD_LOG_FILE
        {
            get
            {
                return this["AUTO_AD_LOG_FILE"] as string;
            }
        }

        [ConfigurationProperty("AD_MAX_REPORT_NUMBER", IsRequired = true)]
        public int AD_MAX_REPORT_NUMBER
        {
            get
            {
                int number = 21;
                int.TryParse(this["AD_MAX_REPORT_NUMBER"].ToString(), out number);
                return number;
            }
        }

        [ConfigurationProperty("BESR_LOG_FILE", IsRequired = true)]
        public string BESR_LOG_FILE
        {
            get
            {
                return this["BESR_LOG_FILE"] as string;
            }
        }

        [ConfigurationProperty("BESR_RESULTS_FOLDER", IsRequired = true)]
        public string BESR_RESULTS_FOLDER
        {
            get
            {
                return this["BESR_RESULTS_FOLDER"] as string;
            }
        }

        [ConfigurationProperty("AUTO_BESR_LOG_FILE", IsRequired = true)]
        public string AUTO_BESR_LOG_FILE
        {
            get
            {
                return this["AUTO_BESR_LOG_FILE"] as string;
            }
        }

        [ConfigurationProperty("BESR_MAX_REPORT_NUMBER", IsRequired = true)]
        public int BESR_MAX_REPORT_NUMBER
        {
            get
            {
                int number = 21;
                int.TryParse(this["BESR_MAX_REPORT_NUMBER"].ToString(), out number);
                return number;
            }
        }

        [ConfigurationProperty("BACKUP_REMOTE_CHECK_SERVER_ROOT_FOLDER", IsRequired = true)]
        public string BACKUP_REMOTE_CHECK_SERVER_ROOT_FOLDER
        {
            get
            {
                return this["BACKUP_REMOTE_CHECK_SERVER_ROOT_FOLDER"] as string;
            }
        }

        [ConfigurationProperty("BACKUP_REMOTE_SERVER_EXEC_ROOT_FOLDER", IsRequired = true)]
        public string BACKUP_REMOTE_SERVER_EXEC_ROOT_FOLDER
        {
            get
            {
                return this["BACKUP_REMOTE_SERVER_EXEC_ROOT_FOLDER"] as string;
            }
        }

        [ConfigurationProperty("BACKUP_REMOTE_SERVER_EXEC", IsRequired = true)]
        public string BACKUP_REMOTE_SERVER_EXEC
        {
            get
            {
                return this["BACKUP_REMOTE_SERVER_EXEC"] as string;
            }
        }

        [ConfigurationProperty("BESR_DEFAULT_INIT_FILE", IsRequired = true)]
        public string BESR_DEFAULT_INIT_FILE
        {
            get
            {
                return this["BESR_DEFAULT_INIT_FILE"] as string;
            }
        }

        [ConfigurationProperty("BESR_RELATIVE_INIT_FILE", IsRequired = true)]
        public string BESR_RELATIVE_INIT_FILE
        {
            get
            {
                return this["BESR_RELATIVE_INIT_FILE"] as string;
            }
        }

        [ConfigurationProperty("BESR_DEFAULT_INIT_FILE_README", IsRequired = true)]
        public string BESR_DEFAULT_INIT_FILE_README
        {
            get
            {
                return this["BESR_DEFAULT_INIT_FILE_README"] as string;
            }
        }

        [ConfigurationProperty("BESR_AUTO_UPDATE_LOG_FILE", IsRequired = true)]
        public string BESR_AUTO_UPDATE_LOG_FILE
        {
            get
            {
                return this["BESR_AUTO_UPDATE_LOG_FILE"] as string;
            }
        }

        [ConfigurationProperty("APP_LOG_FILE", IsRequired = true)]
        public string APP_LOG_FILE
        {
            get
            {
                return this["APP_LOG_FILE"] as string;
            }
        }

        [ConfigurationProperty("APP_RESULTS_FOLDER", IsRequired = true)]
        public string APP_RESULTS_FOLDER
        {
            get
            {
                return this["APP_RESULTS_FOLDER"] as string;
            }
        }

        [ConfigurationProperty("AUTO_APP_LOG_FILE", IsRequired = true)]
        public string AUTO_APP_LOG_FILE
        {
            get
            {
                return this["AUTO_APP_LOG_FILE"] as string;
            }
        }

        [ConfigurationProperty("APP_MAX_REPORT_NUMBER", IsRequired = true)]
        public int APP_MAX_REPORT_NUMBER
        {
            get
            {
                int number = 21;
                int.TryParse(this["APP_MAX_REPORT_NUMBER"].ToString(), out number);
                return number;
            }
        }

        [ConfigurationProperty("APP_DEFAULT_INIT_FILE", IsRequired = true)]
        public string APP_DEFAULT_INIT_FILE
        {
            get
            {
                return this["APP_DEFAULT_INIT_FILE"] as string;
            }
        }

        [ConfigurationProperty("APP_RELATIVE_INIT_FILE", IsRequired = true)]
        public string APP_RELATIVE_INIT_FILE
        {
            get
            {
                return this["APP_RELATIVE_INIT_FILE"] as string;
            }
        }

        [ConfigurationProperty("APP_DEFAULT_INIT_FILE_README", IsRequired = true)]
        public string APP_DEFAULT_INIT_FILE_README
        {
            get
            {
                return this["APP_DEFAULT_INIT_FILE_README"] as string;
            }
        }

        [ConfigurationProperty("SPACE_LOG_FILE", IsRequired = true)]
        public string SPACE_LOG_FILE
        {
            get
            {
                return this["SPACE_LOG_FILE"] as string;
            }
        }

        [ConfigurationProperty("SPACE_RESULTS_FOLDER", IsRequired = true)]
        public string SPACE_RESULTS_FOLDER
        {
            get
            {
                return this["SPACE_RESULTS_FOLDER"] as string;
            }
        }

        [ConfigurationProperty("AUTO_SPACE_LOG_FILE", IsRequired = true)]
        public string AUTO_SPACE_LOG_FILE
        {
            get
            {
                return this["AUTO_SPACE_LOG_FILE"] as string;
            }
        }

        [ConfigurationProperty("SPACE_MAX_REPORT_NUMBER", IsRequired = true)]
        public int SPACE_MAX_REPORT_NUMBER
        {
            get
            {
                int number = 21;
                int.TryParse(this["SPACE_MAX_REPORT_NUMBER"].ToString(), out number);
                return number;
            }
        }

        [ConfigurationProperty("SPACE_DEFAULT_INIT_FILE", IsRequired = true)]
        public string SPACE_DEFAULT_INIT_FILE
        {
            get
            {
                return this["SPACE_DEFAULT_INIT_FILE"] as string;
            }
        }

        [ConfigurationProperty("SPACE_RELATIVE_INIT_FILE", IsRequired = true)]
        public string SPACE_RELATIVE_INIT_FILE
        {
            get
            {
                return this["SPACE_RELATIVE_INIT_FILE"] as string;
            }
        }

        [ConfigurationProperty("SPACE_DEFAULT_INIT_FILE_README", IsRequired = true)]
        public string SPACE_DEFAULT_INIT_FILE_README
        {
            get
            {
                return this["SPACE_DEFAULT_INIT_FILE_README"] as string;
            }
        }

        [ConfigurationProperty("SPACE_DEFAULT_THRESHOLD", IsRequired = true)]
        public double SPACE_DEFAULT_THRESHOLD
        {
            get
            {
                NumberFormatInfo provider = new NumberFormatInfo();
                provider.NumberDecimalSeparator = ".";
                double threshold = 0.1;
                threshold = Convert.ToDouble(this["SPACE_DEFAULT_THRESHOLD"].ToString(), provider);
                return threshold;
            }
        }

        [ConfigurationProperty("DEFAULT_NUMBER_DECIMAL_SEPARATOR", IsRequired = true)]
        public string DEFAULT_NUMBER_DECIMAL_SEPARATOR
        {
            get
            {
                return this["DEFAULT_NUMBER_DECIMAL_SEPARATOR"] as string;
            }
        }

        [ConfigurationProperty("DEFAULT_OCTECT_INCREMENT", IsRequired = true)]
        public double DEFAULT_OCTECT_INCREMENT
        {
            get
            {
                NumberFormatInfo provider = new NumberFormatInfo();
                provider.NumberDecimalSeparator = ".";
                double increment = 1024;
                increment = Convert.ToDouble(this["DEFAULT_OCTECT_INCREMENT"].ToString(), provider);
                return increment;
            }
        }

        [ConfigurationProperty("LOG_SIZE_LIMIT", IsRequired = true)]
        public double LOG_SIZE_LIMIT
        {
            get
            {
                NumberFormatInfo provider = new NumberFormatInfo();
                provider.NumberDecimalSeparator = ".";
                double increment = 100;
                increment = Convert.ToDouble(this["LOG_SIZE_LIMIT"].ToString(), provider);
                return increment;
            }
        }

        [ConfigurationProperty("SPACE_MAX_CHARTS_LINES_NUMBER", IsRequired = true)]
        public int SPACE_MAX_CHARTS_LINES_NUMBER
        {
            get
            {
                int number = 14;
                int.TryParse(this["SPACE_MAX_CHARTS_LINES_NUMBER"].ToString(), out number);
                return number;
            }
        }

        [ConfigurationProperty("SPACE_DEFAULT_LOCAL_MAPPED_DRIVE_LETTER", IsRequired = true)]
        public string SPACE_DEFAULT_LOCAL_MAPPED_DRIVE_LETTER
        {
            get
            {
                return this["SPACE_DEFAULT_LOCAL_MAPPED_DRIVE_LETTER"] as string;
            }
        }

        [ConfigurationProperty("SPACE_USERNAME_IMPERSONNATION", IsRequired = true)]
        public string SPACE_USERNAME_IMPERSONNATION
        {
            get
            {
                return this["SPACE_USERNAME_IMPERSONNATION"] as string;
            }
        }

        [ConfigurationProperty("SPACE_DOMAIN_IMPERSONNATION", IsRequired = true)]
        public string SPACE_DOMAIN_IMPERSONNATION
        {
            get
            {
                return this["SPACE_DOMAIN_IMPERSONNATION"] as string;
            }
        }

        [ConfigurationProperty("SPACE_PASSWORD_IMPERSONNATION", IsRequired = true)]
        public string SPACE_PASSWORD_IMPERSONNATION
        {
            get
            {
                return this["SPACE_PASSWORD_IMPERSONNATION"] as string;
            }
        }

        [ConfigurationProperty("MCO_EASY_TOOL_BASE_URL", IsRequired = true)]
        public string MCO_EASY_TOOL_BASE_URL
        {
            get
            {
                return this["MCO_EASY_TOOL_BASE_URL"] as string;
            }
        }

        [ConfigurationProperty("MCO_EASY_TOOL_SITENAME", IsRequired = true)]
        public string MCO_EASY_TOOL_SITENAME
        {
            get
            {
                return this["MCO_EASY_TOOL_SITENAME"] as string;
            }
        }

    }
}
