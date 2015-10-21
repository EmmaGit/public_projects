using System;
using System.Collections.Generic;
using System.Web.Mvc;


namespace McoEasyTool.Controllers
{
    [AllowAnonymous]
    public class ReportsController : Controller
    {
        private DataModelContainer db = new DataModelContainer();

        public string ReSend(int id)
        {
            Report report = db.Reports.Find(id);
            if (report != null)
            {
                EmailsController emailscontroller = new EmailsController();
                JsonResult JsonSent = emailscontroller.ReSend(report.Email.Id);
                Dictionary<string, string> result = (Dictionary<string, string>)JsonSent.Data;
                McoUtilities.General_Logging(new Exception("...."), "ReSend Report " + User.Identity.Name, 3);
                return result["Response"];
            }
            McoUtilities.General_Logging(new Exception("...."), "ReSend Report", 2, User.Identity.Name);
            return "Le rapport n'a pas été retrouvé dans la base de données.";
        }

        public string ViewEmail(int id)
        {
            Report report = db.Reports.Find(id);
            if (report == null)
            {
                return HttpNotFound().ToString();
            }
            if (report.Email == null)
            {
                return HttpNotFound().ToString();
            }
            Dictionary<string, string> result = new Dictionary<string, string>();
            string window = "<div style='width:100%;height:100%;" +
                            "background-color:white;display:block;font-size-14px;'>" +
                            "<span>Destinataire : " + report.Email.Recipients + "</span><br />" +
                            "<span>Sujet : " + report.Email.Subject + "</span><br />" +
                            "<div>" + report.Email.Body + "</div></div>";
            return window;
        }

        public string Download(int id)
        {
            Report report = db.Reports.Find(id);
            if (report == null)
            {
                return HttpNotFound().ToString();
            }
            try
            {
                System.Web.HttpResponse response = System.Web.HttpContext.Current.Response;
                response.ClearContent();
                response.Clear();
                response.ContentType = "text/plain";
                response.AddHeader("Content-Disposition", "attachment; filename=Rapport_" + report.Module + "_" + report.DateTime.ToString("dd") +
                    "_" + report.DateTime.ToString("MM") + "_" + report.DateTime.ToString("yyyy") + "_" +
                    report.DateTime.ToString("HH") + "h" + report.DateTime.ToString("mm") + "m" +
                    report.ResultPath.Substring(report.ResultPath.LastIndexOf(".")).Trim());
                String RelativePath = report.ResultPath.Replace(Request.ServerVariables["APPL_PHYSICAL_PATH"], String.Empty);
                response.TransmitFile(report.ResultPath);
                response.Flush();
                response.End();
                McoUtilities.General_Logging(new Exception("...."), "Download Report", 3, User.Identity.Name);
                return "OK";
            }
            catch (Exception exception)
            {
                McoUtilities.General_Logging(new Exception("...."), "Download Report");
                return exception.Message;
            }
        }

        public string HardDelete(string file) 
        {
            string result = "";
            try
            {
                System.IO.File.Delete(file);
                result = file + " deleted";
            }
            catch(Exception exception) 
            {
                result = file + " not deleted:\n" + exception.Message;
            }
            return result;
        }

        public string Delete(int id)
        {
            string result = "";
            Report report = db.Reports.Find(id);
            try
            {
                Email email = report.Email;
                string module = report.Module;
                string datetime = report.DateTime.ToString();
                string user = (User != null) ? User.Identity.Name : "N/A";
                string file = report.ResultPath;
                db.Emails.Remove(email);
                switch (module)
                {
                    case HomeController.AD_MODULE:
                        AdReport ad_report = db.AdReports.Find(report.Id);
                        db.AdReports.Remove(ad_report);
                        break;
                    case HomeController.BESR_MODULE:
                        BackupReport besr_report = db.BackupReports.Find(report.Id);
                        db.BackupReports.Remove(besr_report);
                        break;
                    case HomeController.APP_MODULE:
                        AppReport app_report = db.AppReports.Find(report.Id);
                        db.AppReports.Remove(app_report);
                        break;
                    case HomeController.SPACE_MODULE:
                        SpaceReport space_report = db.SpaceReports.Find(report.Id);
                        db.SpaceReports.Remove(space_report);
                        break;
                    default:
                        db.Reports.Remove(report);
                        break;
                }
                db.SaveChanges();
                result = "Le rapport a été correctement supprimé";
                result += "\n" + HardDelete(file);
                McoUtilities.General_Logging(new Exception("...."), "Delete Report by " + User.Identity.Name, 3, User.Identity.Name);
            }
            catch (Exception exception)
            {
                McoUtilities.General_Logging(exception, "Delete Report", 0);
                result = "Une erreur est surveunue lors de la suppression" +
                    exception.Message;
            }
            return result;
        }

        protected override void Dispose(bool disposing)
        {
            db.Dispose();
            base.Dispose(disposing);
        }

    }
}