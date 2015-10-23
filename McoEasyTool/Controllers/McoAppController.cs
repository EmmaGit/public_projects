using NUnit.Framework;
using OpenQA.Selenium;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.Entity;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Management;
using System.Net;
using System.Net.NetworkInformation;
using System.Security.Principal;
using System.ServiceProcess;
using System.Web;
using System.Web.Mvc;
using Excel = Microsoft.Office.Interop.Excel;

namespace McoEasyTool.Controllers
{
    [AllowAnonymous]
    public class McoAppController : Controller
    {
        private DataModelContainer db = new DataModelContainer();
        private ReportsController Reports_Controller = new ReportsController();
        private EmailsController Emails_Controller = new EmailsController();
        private SchedulesController Schedules_Controller = new SchedulesController();
        private ServersController Servers_Controller = new ServersController();
        private McoUtilities.UNCAccessWithCredentials UNC_ACCESSOR = new McoUtilities.UNCAccessWithCredentials();
        private static Excel.Workbook MyWorkbook = null;
        private static Excel.Application MyApplication = null;
        private static Excel.Worksheet MySheet = null;

        public ActionResult Home()
        {
            ViewBag.APP_DESC_0 = McoUtilities.GetModuleDescription(HomeController.APP_MODULE, 0);
            ViewBag.APP_DESC_1 = McoUtilities.GetModuleDescription(HomeController.APP_MODULE, 1);
            ViewBag.APP_DESC_2 = McoUtilities.GetModuleDescription(HomeController.APP_MODULE, 2);
            ViewBag.APP_DESC_3 = McoUtilities.GetModuleDescription(HomeController.APP_MODULE, 3);
            ViewBag.APP_DESC_4 = McoUtilities.GetModuleDescription(HomeController.APP_MODULE, 4);
            return View();
        }

        public ActionResult DisplayRecipients()
        {
            return View(db.Recipients.Where(rec => rec.Module == HomeController.APP_MODULE).ToList());
        }

        public ActionResult UploadInitFile()
        {
            HttpPostedFileBase file = Request.Files[0];
            if (file != null)
            {
                file.SaveAs(HomeController.APP_INIT_FILE);
            }
            ViewBag.Message = Import(true);
            return View();
        }

        public ActionResult DisplayImporter()
        {
            return View();
        }

        public ActionResult DisplayChecker()
        {
            return View(db.Applications.OrderBy(app => app.Name).ToList());
        }

        public ActionResult DisplayFurhterChecker()
        {
            return View(db.Applications.OrderBy(app => app.Name).ToList());
        }

        public ActionResult FunctionLauncher()
        {
            string[] applicationsId;
            string action = "";
            string answer = "";
            string separator = "<table style='position:relative;width:100%;background-color :#aca8a4;'></table>";
            List<Application> SelectedApps = new List<Application>();
            try
            {
                applicationsId = Request.Form["selectedApps"].ToString().Split(',');
                action = Request.Form["action"].ToString();
                foreach (string applicationId in applicationsId)
                {
                    if (applicationId.Trim() == "")
                    {
                        continue;
                    }
                    int appId = 0;
                    Int32.TryParse(applicationId.Trim(), out appId);
                    Application application = db.Applications.Find(appId);
                    if (application == null)
                    {
                        ViewBag.Response = "L'application n'a pas été retrouvée dans la base de données.";
                        return View();
                    }
                    switch (action)
                    {
                        case "CHECK":
                            answer += "<br/ >" + separator + "Vérification " + application.Name + " :";
                            JsonResult CheckJsonResponse = GetApplicationState(appId);
                            Dictionary<string, string> Checkresponse = (Dictionary<string, string>)CheckJsonResponse.Data;
                            answer += Checkresponse["status"] + "<br/ >";
                            answer += Checkresponse["details"] + "<br/ >";
                            break;
                        case "START":
                            answer += "<br/ >" + separator + "Démarrage " + application.Name + " :";
                            JsonResult StartJsonResponse = StartApplication(appId);
                            Dictionary<string, string> Startresponse = (Dictionary<string, string>)StartJsonResponse.Data;
                            answer += Startresponse["status"] + "<br/ >";
                            answer += Startresponse["details"] + "<br/ >";
                            break;
                        case "RESTART":
                            answer += "<br/ >" + separator + "Redémarrage " + application.Name + " :";
                            JsonResult RestartJsonResponse = RestartApplication(appId);
                            Dictionary<string, string> Restartresponse = (Dictionary<string, string>)RestartJsonResponse.Data;
                            answer += Restartresponse["status"] + "<br/ >";
                            answer += Restartresponse["details"] + "<br/ >";
                            break;
                        case "STOP":
                            answer += "<br/ >" + separator + "Arrêt " + application.Name + " :";
                            JsonResult StopJsonResponse = StopApplication(appId);
                            Dictionary<string, string> Stopresponse = (Dictionary<string, string>)StopJsonResponse.Data;
                            answer += Stopresponse["status"] + "<br/ >";
                            answer += Stopresponse["details"] + "<br/ >";
                            break;
                        case "DELETE":
                            answer += "<br/ >" + separator + "Suppression " + application.Name + " :<br/ >" +
                                DeleteApplication(appId) + "<br/ >";
                            break;
                    }
                }
            }
            catch (Exception exception)
            {
                string log_path_info = exception.Message + "\r\n";
                try
                {
                    string log = "\r\n**************************************************\r\n";
                    log += DateTime.Now.ToString() + " : " + "FunctionLauncher Error : General Error \r\n";
                    log += exception.Message + "\r\n";
                    System.IO.File.AppendAllText(HomeController.GENERAL_LOG_FILE, log);
                }
                catch { }
                ViewBag.Response = "Une exception est survenue: " + exception.Message;
                return View();
            }
            ViewBag.Response = answer;
            return View();
        }

        public ActionResult DisplayAppDomains()
        {
            AppDomain[] domains = db.AppDomains.ToArray();
            foreach (AppDomain domain in domains)
            {
                domain.Applications = db.Applications.Where(app => app.Domain == domain.Id).Count();
                db.Entry(domain).State = System.Data.Entity.EntityState.Modified;
                db.SaveChanges();
            }
            return View(db.AppDomains.OrderBy(domain => domain.Name).ToList());
        }

        public ActionResult DisplayScheduledApplicationsList(int id)
        {
            AppSchedule schedule = db.AppSchedules.Find(id);
            if (schedule == null)
            {
                return HttpNotFound();
            }
            object[] boundaries =
                McoUtilities.GetIdValues<AppSchedule>(schedule, HomeController.OBJECT_ATTR_ID);
            if (boundaries != null)
            {
                ViewBag.First_Entry = boundaries[0].ToString();
                ViewBag.Left_Entry = boundaries[1].ToString();
                ViewBag.Current_Entry = boundaries[2].ToString();
                ViewBag.Right_Entry = boundaries[3].ToString();
                ViewBag.Last_Entry = boundaries[4].ToString();
            }
            else
            {
                ViewBag.First_Entry = "0";
                ViewBag.Left_Entry = "0";
                ViewBag.Current_Entry = "0";
                ViewBag.Right_Entry = "0";
                ViewBag.Last_Entry = "0";
            }

            ViewBag.ScheduleId = id;
            ViewBag.Message = schedule.TaskName;
            return View(db.Applications.ToList());
        }

        public ActionResult DisplayScheduleReports(int id)
        {
            AppSchedule schedule = db.AppSchedules.Find(id);
            object[] boundaries =
                McoUtilities.GetIdValues<AppSchedule>(schedule, HomeController.OBJECT_ATTR_ID);
            if (boundaries != null)
            {
                ViewBag.First_Entry = boundaries[0].ToString();
                ViewBag.Left_Entry = boundaries[1].ToString();
                ViewBag.Current_Entry = boundaries[2].ToString();
                ViewBag.Right_Entry = boundaries[3].ToString();
                ViewBag.Last_Entry = boundaries[4].ToString();
            }
            else
            {
                ViewBag.First_Entry = "0";
                ViewBag.Left_Entry = "0";
                ViewBag.Current_Entry = "0";
                ViewBag.Right_Entry = "0";
                ViewBag.Last_Entry = "0";
            }
            ViewBag.Message = "Rapports du check planifié " + schedule.TaskName;
            ICollection<AppReport> reports = db.AppReports.Where(report => report.ScheduleId == id).ToList();
            return View(reports.OrderByDescending(report => report.Id).ToList());
        }

        public ActionResult DisplayScheduleReportDetails(int id)
        {
            AppReport report = db.AppReports.Find(id);
            if (report == null)
            {
                return HttpNotFound();
            }
            ViewBag.Message = report.DateTime.ToString();
            return View(db.ApplicationReports.Where(rep => rep.AppReportId == id).ToList());
        }

        public ActionResult DisplaySchedules()
        {
            return View(db.AppSchedules.OrderBy(name => name.TaskName).ToList());
        }

        public ActionResult DisplayApplications()
        {
            return View(db.Applications.OrderBy(app => app.Name).ToList());
        }

        public ActionResult DisplayAuthenticationInfos(int id)
        {
            Application application = db.Applications.Find(id);
            if (application == null)
            {
                return HttpNotFound();
            }
            object[] boundaries =
                McoUtilities.GetIdValues<Application>(application, HomeController.OBJECT_ATTR_NAME);

            ViewBag.Message = "Procédure d'authentification " + application.Name;
            ViewBag.appId = application.Id;
            ViewBag.appUrl = application.Url;

            if (boundaries != null)
            {
                ViewBag.First_Entry = boundaries[0].ToString();
                ViewBag.Left_Entry = boundaries[1].ToString();
                ViewBag.Current_Entry = boundaries[2].ToString();
                ViewBag.Right_Entry = boundaries[3].ToString();
                ViewBag.Last_Entry = boundaries[4].ToString();
            }
            else
            {
                ViewBag.First_Entry = "0";
                ViewBag.Left_Entry = "0";
                ViewBag.Current_Entry = "0";
                ViewBag.Right_Entry = "0";
                ViewBag.Last_Entry = "0";
            }


            ViewBag.Text_Tags = ViewBag.Src_Tags = ViewBag.Val_Tags = "";
            foreach (string tag in HomeController.APP_TEXT_ATTR_TAGS_LIST)
            {
                ViewBag.Text_Tags += tag + ";";
            }
            foreach (string tag in HomeController.APP_SRC_ATTR_TAGS_LIST)
            {
                ViewBag.Src_Tags += tag + ";";
            }
            foreach (string tag in HomeController.APP_VALUE_ATTR_TAGS_LIST)
            {
                ViewBag.Val_Tags += tag + ";";
            }
            return View(application.AppHtmlElements.ToList());
        }

        public ActionResult DisplayApplicationServers(int id)
        {
            Application application = db.Applications.Find(id);
            object[] boundaries =
                McoUtilities.GetIdValues<Application>(application, HomeController.OBJECT_ATTR_NAME);
            if (boundaries != null)
            {
                ViewBag.First_Entry = boundaries[0].ToString();
                ViewBag.Left_Entry = boundaries[1].ToString();
                ViewBag.Current_Entry = boundaries[2].ToString();
                ViewBag.Right_Entry = boundaries[3].ToString();
                ViewBag.Last_Entry = boundaries[4].ToString();
            }
            else
            {
                ViewBag.First_Entry = "0";
                ViewBag.Left_Entry = "0";
                ViewBag.Current_Entry = "0";
                ViewBag.Right_Entry = "0";
                ViewBag.Last_Entry = "0";
            }
            ViewBag.Message = "Gestion de l'application " + application.Name;
            ViewBag.appId = application.Id;
            return View(application.AppServers.ToList());
        }

        public ActionResult DisplayFailedApplications()
        {
            List<Application_Report> appreports = null;
            AppReport report = db.AppReports.OrderByDescending(rep => rep.Id).First();
            appreports = db.ApplicationReports
                    .Where(state => state.State != "OK")
                    .Where(rep => rep.AppReportId == report.Id)
                    .Where(ser => ser.AppServer_Reports.Count != 0)
                    .OrderBy(name => name.Application.Name)
                    .ToList();
            //ApplicationApp
            return View(appreports);
        }

        public ActionResult DisplayReports()
        {
            return View(db.AppReports.OrderByDescending(id => id.Id).ToList());
        }

        public ActionResult DisplayReportDetails(int id)
        {
            AppReport report = db.AppReports.Find(id);
            if (report == null)
            {
                return HttpNotFound();
            }
            object[] boundaries =
                McoUtilities.GetIdValues<AppReport>(report, HomeController.OBJECT_ATTR_ID, true);
            if (boundaries != null)
            {
                ViewBag.First_Entry = boundaries[0].ToString();
                ViewBag.Left_Entry = boundaries[1].ToString();
                ViewBag.Current_Entry = boundaries[2].ToString();
                ViewBag.Right_Entry = boundaries[3].ToString();
                ViewBag.Last_Entry = boundaries[4].ToString();
            }
            else
            {
                ViewBag.First_Entry = "0";
                ViewBag.Left_Entry = "0";
                ViewBag.Current_Entry = "0";
                ViewBag.Right_Entry = "0";
                ViewBag.Last_Entry = "0";
            }
            ViewBag.Message = report.DateTime.ToString();
            return View(db.ApplicationReports.Where(rep => rep.AppReportId == id).ToList());
        }

        public ActionResult DisplayReportFurtherDetails(int id)
        {
            Application_Report appreport = db.ApplicationReports.Find(id);
            if (appreport == null)
            {
                return HttpNotFound();
            }
            string requested = "";
            try
            {
                requested = Request.Form["origin"].ToString();
            }
            catch
            {
                requested = "";
            }
            if (requested != "")
            {
                ViewBag.Report = requested + "/" + appreport.AppReport.Id.ToString();
            }
            else
            {
                ViewBag.Report = "DisplayReportDetails/" + appreport.AppReport.Id.ToString();
            }
            ViewBag.ApplicationName = appreport.Application.Name;
            ViewBag.Date = appreport.AppServer_Reports.FirstOrDefault().Application_Report.AppReport.DateTime;
            return View(appreport.AppServer_Reports.ToList());
        }

        //ACTIONS
        public string CreateSchedule()
        {
            int day = 0; int month = 0; int year = 0; int hours = 0; int minutes = 0;
            string postedtaskname = Request.Form["taskname"].ToString();
            bool dayOk = Int32.TryParse(Request.Form["day"].ToString(), out day);
            bool monthOk = Int32.TryParse(Request.Form["month"].ToString(), out month);
            bool yearOk = Int32.TryParse(Request.Form["year"].ToString(), out year);
            bool hoursOk = Int32.TryParse(Request.Form["hours"].ToString(), out hours);
            bool minutesOk = Int32.TryParse(Request.Form["minutes"].ToString(), out minutes);
            string multiplicity = Request.Form["multiplicity"].ToString();

            string result = "";
            if (dayOk && monthOk && yearOk && hoursOk && minutesOk)
            {
                TimeSpan time = new TimeSpan(hours, minutes, 30);
                DateTime now = DateTime.Now;
                DateTime scheduled = new DateTime(year, month + 1, day) + time;
                int scheduleId = 0;
                if (scheduled.CompareTo(DateTime.Now) > 0)
                {
                    string taskname = HomeController.APP_MODULE + " AutoCheck ";

                    AppSchedule schedule = db.AppSchedules.Create();
                    schedule.CreationTime = DateTime.Now;
                    schedule.NextExecution = scheduled;
                    schedule.Generator = User.Identity.Name;
                    schedule.Multiplicity = multiplicity;
                    schedule.Executed = 0;
                    schedule.State = "Planifié";
                    schedule.Module = HomeController.APP_MODULE;
                    schedule.AutoRelaunch = false;
                    if (ModelState.IsValid)
                    {
                        db.Schedules.Add(schedule);
                        db.SaveChanges();
                        scheduleId = schedule.Id;
                        taskname += scheduleId;
                        if (postedtaskname != "" && postedtaskname != " " && postedtaskname != null)
                        {
                            schedule.TaskName = postedtaskname;
                        }
                        else
                        {
                            schedule.TaskName = taskname;
                        }
                        db.Entry(schedule).State = System.Data.Entity.EntityState.Modified;
                        db.SaveChanges();
                        return Schedules_Controller.Create(schedule);
                    }
                }
                else
                {
                    result = "Erreur, la date est dépassée.";
                }
            }
            else
            {
                result = "Impossible de programmer le lancement à la date spécifiée. Veuillez réessayer ultérieurement.";
            }
            Specific_Logging(new Exception(""), "CreateSchedule", 3);
            return result;
        }

        public string EditSchedule(int id)
        {
            AppSchedule schedule = db.AppSchedules.Find(id);
            if (schedule == null)
            {
                return HttpNotFound().ToString();
            }

            int day = 0; int month = 0; int year = 0; int hours = 0; int minutes = 0;
            string postedtaskname = Request.Form["taskname"].ToString();
            bool dayOk = Int32.TryParse(Request.Form["day"].ToString(), out day);
            bool monthOk = Int32.TryParse(Request.Form["month"].ToString(), out month);
            bool yearOk = Int32.TryParse(Request.Form["year"].ToString(), out year);
            bool hoursOk = Int32.TryParse(Request.Form["hours"].ToString(), out hours);
            bool minutesOk = Int32.TryParse(Request.Form["minutes"].ToString(), out minutes);
            string multiplicity = Request.Form["multiplicity"].ToString();

            string result = "";
            if (dayOk && monthOk && yearOk && hoursOk && minutesOk)
            {
                TimeSpan time = new TimeSpan(hours, minutes, 30);
                DateTime now = DateTime.Now;
                DateTime scheduled = new DateTime(year, month + 1, day) + time;
                if (scheduled.CompareTo(DateTime.Now) > 0 &&
                    (Schedules_Controller.Delete(schedule) == "La tâche a été correctement supprimée"))
                {
                    string taskname = HomeController.APP_MODULE + " AutoCheck ";

                    schedule.NextExecution = scheduled;
                    schedule.Generator = User.Identity.Name;
                    schedule.Multiplicity = multiplicity;
                    schedule.State = "Planifié";
                    schedule.AutoRelaunch = false;
                    if (ModelState.IsValid)
                    {
                        taskname += id;
                        if (postedtaskname != "" && postedtaskname != " " && postedtaskname != null)
                        {
                            schedule.TaskName = postedtaskname;
                        }
                        else
                        {
                            schedule.TaskName = taskname;
                        }
                        db.Entry(schedule).State = System.Data.Entity.EntityState.Modified;
                        db.SaveChanges();
                        return Schedules_Controller.Edit(schedule);
                    }
                }
                else
                {
                    result = "Erreur, la date est dépassée.";
                }
            }
            else
            {
                result = "Impossible de programmer le lancement à la date spécifiée. Veuillez réessayer ultérieurement.";
            }
            Specific_Logging(new Exception(""), "EditSchedule", 3);
            return result;
        }

        public string DeleteSchedule(int id)
        {
            AppSchedule schedule = db.AppSchedules.Find(id);
            if (schedule.Scheduled_Applications.Count != 0)
            {
                List<Scheduled_Application> apps = schedule.Scheduled_Applications.ToList();
                foreach (Scheduled_Application app in apps)
                {
                    db.ScheduledApplications.Remove(app);
                    db.Entry(schedule).State = System.Data.Entity.EntityState.Modified;
                    db.SaveChanges();
                }
            }
            if (schedule.Reports.Count != 0)
            {
                List<AppReport> reports = db.AppReports.Where(rep => rep.ScheduleId == schedule.Id).ToList();
                foreach (Report report in reports)
                {
                    DeleteAppReport(report.Id);
                    db.Entry(schedule).State = System.Data.Entity.EntityState.Modified;
                    db.SaveChanges();
                }
            }

            string result = Schedules_Controller.Delete(schedule);
            schedule = db.AppSchedules.Find(id);
            if (result == "La tâche a été correctement supprimée")
            {
                db.AppSchedules.Remove(schedule);
                db.SaveChanges();
            }
            Specific_Logging(new Exception(""), "DeleteSchedule", 3);
            return result;
        }

        public string ReSendLastEmail(int id)
        {
            AppSchedule schedule = db.AppSchedules.Find(id);
            if (schedule != null)
            {
                if (db.Reports.Where(report => report.ScheduleId == id).Count() != 0)
                {
                    AppReport report = db.AppReports.Where(rep => rep.ScheduleId == id).OrderByDescending(rep => rep.Id).First();
                    return Reports_Controller.ReSend(report.Id);
                }
                return "Cette tâche planifiée n'a pour l'instant généré aucun rapport, ou alors ils ont été supprimés.";

            }
            Specific_Logging(new Exception(""), "ReSendLastEmail", 3);
            return "Cette tâche planifiée n'a pas été trouvée dans la base de données.";
        }

        public string SaveList(int id)
        {
            AppSchedule appSchedule = db.AppSchedules.Find(id);
            if (appSchedule == null)
            {
                return HttpNotFound().ToString();
            }
            string[] applicationList;
            string applications = "";
            List<Application> SelectedApps = new List<Application>();
            try
            {
                applications = Request.Form["list"].ToString();
                applicationList = applications.Split(';');
                foreach (string appId in applicationList)
                {
                    int schappid = 0;
                    Int32.TryParse(appId.Split('-')[1], out schappid);
                    Application application = db.Applications.Find(schappid);
                    if (application != null)
                    {
                        SelectedApps.Add(application);
                    }
                }
                List<Application> ExistingApps = appSchedule.Scheduled_Applications.Select(appid => appid.Application).ToList();
                List<Application> ExcludedApps = ExistingApps.Except(SelectedApps).ToList();
                List<Application> NewApps = SelectedApps.Where(app => !ExistingApps.Contains(app)).ToList();
                foreach (Application application in ExcludedApps)
                {
                    Scheduled_Application scheduledApp = appSchedule.Scheduled_Applications
                        .Where(appid => appid.ApplicationId == application.Id).FirstOrDefault();
                    db.ScheduledApplications.Remove(scheduledApp);
                }
                db.SaveChanges();
                foreach (Application application in NewApps)
                {
                    Scheduled_Application scheduledApp = db.ScheduledApplications.Create();
                    scheduledApp.ApplicationId = application.Id;
                    scheduledApp.Application = application;
                    scheduledApp.AppSchedule = appSchedule;
                    scheduledApp.AppScheduleId = appSchedule.Id;
                    if (ModelState.IsValid)
                    {
                        db.ScheduledApplications.Add(scheduledApp);
                        db.SaveChanges();
                    }
                }
                Specific_Logging(new Exception(""), "SaveList", 3);
                return "Les modifications ont été effectuées à la liste des applications";
            }
            catch (Exception exception)
            {
                Specific_Logging(exception, "SaveList");
                return "Une erreur est survenue lors de la sélection des Applications:\n" + exception.Message;
            }
        }

        public string GetApplicationsList(int id)
        {
            AppSchedule appSchedule = db.AppSchedules.Find(id);
            if (appSchedule == null)
            {
                return HttpNotFound().ToString();
            }
            string applications = "";
            List<Scheduled_Application> scheduledApps = appSchedule.Scheduled_Applications.ToList();
            foreach (Scheduled_Application scheduledApp in scheduledApps)
            {
                applications += scheduledApp.Application.Id + ";";
            }
            if (applications.Length > 0)
            {
                applications.Substring(0, applications.Length - 1);
            }
            return applications;
        }

        [HttpPost]
        public JsonResult AddAppHtmlElement(int id)
        {
            Application application = db.Applications.Find(id);
            if (application == null)
            {
                return Json(HttpNotFound().ToString(), JsonRequestBehavior.AllowGet);
            }
            string Tag_TagName = "", Tag_Xpath = "", Tag_Id = "", Tag_Name = "", Tag_Class = "", Tag_Value = "", Tag_Type = "";
            Dictionary<string, string> response = new Dictionary<string, string>();
            response.Add("status", "");
            response.Add("appId", id.ToString());
            try
            {
                Tag_TagName = Request.Form["tagname"].ToString();
                Tag_Xpath = Request.Form["tag_xpath"].ToString();
                Tag_Id = Request.Form["tag_id"].ToString();
                Tag_Name = Request.Form["tag_name"].ToString();
                Tag_Class = Request.Form["tag_class"].ToString();
                Tag_Value = Request.Form["tag_value"].ToString();
                Tag_Type = Request.Form["tag_type"].ToString();

                AppHtmlElement htmlelement = db.AppHtmlElements.Create();
                htmlelement.TagName = Tag_TagName; htmlelement.Type = Tag_Type;
                htmlelement.AttrXpath = Tag_Xpath;
                htmlelement.AttrId = Tag_Id; htmlelement.AttrName = Tag_Name;
                htmlelement.AttrClass = Tag_Class; htmlelement.Value = Tag_Value;
                htmlelement.ApplicationId = application.Id;

                AppHtmlElement login = application.AppHtmlElements.Where(el => el.Type == "LOGIN").FirstOrDefault();
                if (ModelState.IsValid)
                {
                    if (login != null && htmlelement.Type == "LOGIN")
                    {
                        response["status"] = "KO: Il y a déjà un bouton de connexion";
                        return Json(response, JsonRequestBehavior.AllowGet);
                    }
                    db.AppHtmlElements.Add(htmlelement);
                    db.SaveChanges();
                    response["status"] = "OK";
                    response["appId"] = application.Id.ToString();
                    Specific_Logging(new Exception(""), "AddAppHtmlElement", 3);
                    return Json(response, JsonRequestBehavior.AllowGet);
                }
                response["status"] = "KO";
                Specific_Logging(new Exception(""), "AddAppHtmlElement", 2);
                return Json(response, JsonRequestBehavior.AllowGet);
            }
            catch (Exception exception)
            {
                Specific_Logging(exception, "AddAppHtmlElement", 3);
                response["status"] = "Erreur lors de l'ajout du paramètre";
                return Json(response, JsonRequestBehavior.AllowGet);
            }
        }

        [HttpPost]
        public JsonResult EditAppHtmlElement(int id)
        {
            AppHtmlElement htmlelement = db.AppHtmlElements.Find(id);
            if (htmlelement == null)
            {
                return Json(HttpNotFound().ToString(), JsonRequestBehavior.AllowGet);
            }
            string Tag_TagName = "", Tag_Xpath = "", Tag_Id = "", Tag_Name = "", Tag_Class = "", Tag_Value = "", Tag_Type = "";
            Dictionary<string, string> response = new Dictionary<string, string>();
            response.Add("status", "");
            response.Add("appId", id.ToString());
            try
            {
                Tag_TagName = Request.Form["tagname"].ToString();
                Tag_Xpath = Request.Form["tag_xpath"].ToString();
                Tag_Id = Request.Form["tag_id"].ToString();
                Tag_Name = Request.Form["tag_name"].ToString();
                Tag_Class = Request.Form["tag_class"].ToString();
                Tag_Value = Request.Form["tag_value"].ToString();
                Tag_Type = Request.Form["tag_type"].ToString();

                htmlelement.TagName = Tag_TagName; htmlelement.Type = Tag_Type;
                htmlelement.AttrId = Tag_Id; htmlelement.AttrName = Tag_Name;
                htmlelement.AttrClass = Tag_Class; htmlelement.Value = Tag_Value;
                htmlelement.AttrXpath = Tag_Xpath;

                if (ModelState.IsValid)
                {
                    db.Entry(htmlelement).State = System.Data.Entity.EntityState.Modified;
                    db.SaveChanges();
                    response["status"] = "OK";
                    response["appId"] = htmlelement.Application.Id.ToString();
                    Specific_Logging(new Exception(""), "EditAppHtmlElement", 3);
                    return Json(response, JsonRequestBehavior.AllowGet);
                }
                Specific_Logging(new Exception(""), "EditAppHtmlElement", 2);
                response["status"] = "KO";
                return Json(response, JsonRequestBehavior.AllowGet);
            }
            catch (Exception exception)
            {
                Specific_Logging(exception, "EditAppHtmlElement");
                response["status"] = "Erreur lors de l'ajout du paramètre";
                return Json(response, JsonRequestBehavior.AllowGet);
            }
        }

        public string DeleteAppHtmlElement(int id)
        {
            AppHtmlElement htmlelement = db.AppHtmlElements.Find(id);
            if (htmlelement == null)
            {
                return HttpNotFound().ToString();
            }
            db.AppHtmlElements.Remove(htmlelement);
            db.SaveChanges();
            Specific_Logging(new Exception(""), "DeleteAppHtmlElement", 3);
            return "Le paramètre a été supprimé";
        }

        public string BuildEmail(int id)
        {
            Email email = db.Emails.Find(id);
            if (email == null)
            {
                return HttpNotFound().ToString();
            }

            string body = "<style>tr:hover>td,tr:hover>td a{cursor:pointer;color:#fff;background-color:#68b3ff;}</style>";

            Application_Report[] applicationreports = db.ApplicationReports.Where(
                apprep => apprep.AppReportId == email.Report.Id).ToArray();
            List<Application> applications = new List<Application>();
            foreach (Application_Report applicationreport in applicationreports)
            {
                if (!applications.Contains(applicationreport.Application))
                {
                    applications.Add(applicationreport.Application);
                }
            }


            List<AppDomain> domains = new List<AppDomain>();
            foreach (Application application in applications)
            {
                AppDomain dom = db.AppDomains.Find(application.Domain);
                if (!domains.Contains(dom))
                {
                    domains.Add(dom);
                }
            }
            body += "<br /><table style='position:relative;width:100%;' cellpadding='0' cellspacing='0'>" +
                "<thead><tr style='position:relative;text-align:center;width:100%;background-color:#dcdbdb;'>" +
                    "<th style='position:relative;text-align:center;font-weight:bold;width:8%;border:1px solid #fff''>Domaines</th>" +
                    "<th style='position:relative;text-align:center;width:10%;font-weight:bold;border:1px solid #fff'>Applications</th>" +
                    "<th style='position:relative;text-align:center;width:10%;font-weight:bold;border:1px solid #fff'>Serveurs</th>" +
                    "<th style='position:relative;text-align:center;width:5%;font-weight:bold;border:1px solid #fff'>Etat</th>" +
                    "<th style='position:relative;text-align:center;width:7%;font-weight:bold;border:1px solid #fff'>Ping</th>" +
                    "<th style='position:relative;text-align:center;width:30%;font-weight:bold;border:1px solid #fff'>Details</th>" +
                    "<th style='position:relative;text-align:center;width:30%;font-weight:bold;border:1px solid #fff'>Authentification</th>" +
                "</thead><tbody>";


            string empty = "<span style='white-space:nowrap;height:100px;'></span>";
            foreach (AppDomain domain in domains.OrderBy(dom => dom.Name))
            {
                Application_Report[] dom_applicationreports = applicationreports.Where(app => app.Application.Domain == domain.Id).ToArray();
                int dom_rowspan = 0;
                foreach (Application_Report applicationreport in dom_applicationreports.OrderBy(app => app.Application.Name))
                {
                    int app_rowspan = (applicationreport.AppServer_Reports.Count == 0) ? 1 : applicationreport.AppServer_Reports.Count;
                    dom_rowspan += app_rowspan;
                }
                body += "<tr style='position:relative;text-align:center;font-weight:bold;width:100%;'>" +
                    "<td rowspan='" + dom_rowspan + "'  style='position:relative;text-align:center;border:1px solid #fff;background-color:#dcdbdb;'>" + domain.Name + "</td>";

                foreach (Application_Report applicationreport in dom_applicationreports.OrderBy(app => app.Application.Name))
                {
                    string applicationbgcolor = (applicationreport.State == "OK") ? "#22b14c" :
                       (applicationreport.State == "KO") ? "#ff3f3f" :
                       (applicationreport.State == "H-OK") ? "#de5a26" : "#eeece1";
                    string AuthColor = (applicationreport.Authentified == "OK") ? "#22b14c" : "#ff2f00";
                    string auth = (applicationreport.Authentified.Length < 10) ? "Authentification : " + applicationreport.Authentified : applicationreport.Authentified;

                    int app_rowspan = (applicationreport.AppServer_Reports.Count == 0) ? 1 : applicationreport.AppServer_Reports.Count;
                    if (applicationreport == dom_applicationreports.OrderBy(app => app.Application.Name).First())
                    {
                        body += "<td rowspan='" + app_rowspan + "' style='position:relative;text-align:center;border:1px solid #dcdbdb;background-color:" + applicationbgcolor + "'>" +
                         "<a style='white-space:nowrap;height:100px;width:100%' href='" + applicationreport.Application.Url + "'>" + applicationreport.Application.Name + "</a></td>";
                    }
                    else
                    {
                        body += "<tr style='position:relative;text-align:center;font-weight:bold;width:100%;'>" +
                            "<td rowspan='" + app_rowspan + "' style='position:relative;text-align:center;border:1px solid #dcdbdb;background-color:" + applicationbgcolor + "'>" +
                            "<a href='" + applicationreport.Application.Url + "'>" + applicationreport.Application.Name + "</a></td>";
                    }

                    if (applicationreport.AppServer_Reports.Count == 0)
                    {
                        body += "<td style='position:relative;text-align:center;border:1px solid #dcdbdb;background-color:#eeece1;'>" + empty + "</td>" +
                            "<td style='position:relative;text-align:center;border:1px solid #dcdbdb;background-color:#eeece1;'>" + empty + "</td>" +
                            "<td style='position:relative;text-align:center;border:1px solid #dcdbdb;background-color:#eeece1;'>" + empty + "</td>" +
                            "<td style='position:relative;text-align:center;border:1px solid #dcdbdb;background-color:#eeece1;'>" + empty + "</td>" +
                            "<td style='position:relative;text-align:center;border:1px solid #dcdbdb;background-color:#fff;color:" + AuthColor + "'>" + auth + "</td>";
                        body += "</tr>";
                    }
                    else
                    {
                        AppServer_Report[] serverreports = applicationreport.AppServer_Reports.Where(
                            serverreportid => serverreportid.Application_Report.AppReport.Id == email.Report.Id).ToArray();
                        foreach (AppServer_Report serverreport in serverreports)
                        {
                            string serverbgcolor = (serverreport.State == "OK") ? "#22b14c" : "#ff3f3f";
                            string pingcolor = (serverreport.Ping == "OK") ? "#22b14c" : "#ff3f3f";
                            if (serverreport == serverreports.First())
                            {
                                body += "<td style='position:relative;margin:0px;padding:0px;border:1px solid #dcdbdb;background-color:" + serverbgcolor + ";'>" + serverreport.AppServer.Name + "</td>" +
                                "<td style='position:relative;margin:0px;padding:0px;border:1px solid #dcdbdb;background-color:" + serverbgcolor + ";'>" + serverreport.State + "</td>" +
                                "<td style='position:relative;margin:0px;padding:0px;border:1px solid #dcdbdb;background-color:#fff;color:" + pingcolor + ";'>Ping " + serverreport.Ping + "</td>" +
                                "<td style='position:relative;margin:0px;padding:0px;border:1px solid #dcdbdb;background-color:" + serverbgcolor + ";'>" + serverreport.Details + "</td>" +
                                "<td rowspan='" + app_rowspan + "' style='position:relative;text-align:center;border:1px solid #dcdbdb;background-color:#fff;color:" + AuthColor + "'>" + auth + "</td></tr>";
                            }
                            else
                            {
                                body += "<tr style='position:relative;text-align:center;font-weight:bold;width:100%;'>" +
                                    "<td style='position:relative;margin:0px;padding:0px;border:1px solid #dcdbdb;background-color:" + serverbgcolor + ";'>" + serverreport.AppServer.Name + "</td>" +
                                    "<td style='position:relative;margin:0px;padding:0px;border:1px solid #dcdbdb;background-color:" + serverbgcolor + ";'>" + serverreport.State + "</td>" +
                                    "<td style='position:relative;margin:0px;padding:0px;border:1px solid #dcdbdb;background-color:#fff;color:" + pingcolor + ";'>Ping " + serverreport.Ping + "</td>" +
                                    "<td style='position:relative;margin:0px;padding:0px;border:1px solid #dcdbdb;background-color:" + serverbgcolor + ";'>" + serverreport.Details + "</td></tr>";
                            }
                        }
                    }
                }
            }
            body += "</tbody></table>";
            body += "<br /><table style='position:relative;width:100%;background-color :#aca8a4;'></table>";
            body += "<div>RECAPITULATIF DU RAPPORT : ";
            body += "Date de génération : " + DateTime.Now.ToString() + "<br />";
            body += "<table style='position:relative;width:100%;background-color :#aca8a4;'></table>" +
                "Nombre total d'applications vérifiées : " + email.Report.TotalChecked + "; " +
                "Nombre total d'applications KO : " + email.Report.TotalErrors + " | ";
            double percentage = Math.Round((double)email.Report.TotalErrors / (double)email.Report.TotalChecked, 4) * 100;
            body += "<span style='font-weight:bold;color:#e75114;'>Pourcentage d'applications KO : " + percentage + "%</span><br />";
            body += "<span style='font-weight:bold;color:#e75114;'>Les lignes rouges font références aux serveurs d'applications pour lesquels des erreurs ont été remontées.</span><br />" +
                "<span style='font-weight:bold;color:#e75114;'>Les lignes vertes à l'inverse des rouges n'ont remontées aucune erreur.</span><br />" +
                "<span style='font-weight:bold;color:#e75114;'>Les lignes grises désignent les applications pour lesquelles aucun serveur n'a été renseigné.</span><br />";

            email.Body = body;
            email.Subject = "Resultat check d'Applications " + email.Report.DateTime.ToString();
            if (ModelState.IsValid)
            {
                db.Entry(email).State = System.Data.Entity.EntityState.Modified;
                db.SaveChanges();
            }
            Specific_Logging(new Exception(""), "BuildEmail", 3);
            return "BuildOK";
        }

        public string DeleteAppDomain(int id = 0)
        {
            AppDomain appdomain = db.AppDomains.Find(id);
            if (appdomain == null)
            {
                return HttpNotFound().ToString();
            }
            if (db.Applications.Where(dom => dom.Domain == appdomain.Id).Count() == 0)
            {
                db.AppDomains.Remove(appdomain);
                db.SaveChanges();
                Specific_Logging(new Exception(""), "DeleteAppDomain", 3);
                return "Le domaine a été supprimé avec succès";
            }
            else
            {
                string list = "";
                foreach (Application application in db.Applications.Where(dom => dom.Domain == appdomain.Id))
                {
                    list += application.Name + "\n";
                }
                return "Veuillez d'abord supprimer les applications suivantes:\n" + list;
            }
        }

        public string AddAppDomain()
        {
            try
            {
                string name = Request.Form["name"];
                AppDomain domain = db.AppDomains.Create();
                domain.Name = name;
                domain.Applications = 0;
                if (ModelState.IsValid)
                {
                    db.AppDomains.Add(domain);
                    db.SaveChanges();
                    Specific_Logging(new Exception("...."), "AddAppDomain " + domain.Name, 3);
                    return "Le domaine a été correctement rajouté à la liste.";
                }
                Specific_Logging(new Exception(""), "AddAppDomain", 2);
                return "Erreur lors de l'ajout du domaine";
            }
            catch (Exception exception)
            {
                Specific_Logging(exception, "EditAppHtmlElement");
                return "Erreur lors de l'ajout du domaine";
            }
        }

        public string EditAppDomain(int id)
        {
            try
            {
                string name = Request.Form["name"];
                AppDomain domain = db.AppDomains.Find(id);
                if (domain != null)
                {
                    domain.Name = name;
                    if (ModelState.IsValid)
                    {
                        db.Entry(domain).State = System.Data.Entity.EntityState.Modified;
                        db.SaveChanges();
                        Specific_Logging(new Exception("...."), "EditAppDomain " + domain.Name, 3); ;
                        return "Le domaine a été correctement modifié.";
                    }
                    Specific_Logging(new Exception(""), "EditAppDomain " + domain.Name, 2);
                    return "Erreur lors de la modification du domaine";
                }
                else
                {
                    return "Erreur domaine non trouvé";
                }
            }
            catch (Exception exception)
            {
                Specific_Logging(new Exception(""), "EditAppDomain");
                return "Erreur lors de la modification du domaine";
            }
        }

        public JsonResult GetAppDomainsList()
        {
            AppDomain[] domains = db.AppDomains.OrderBy(dom => dom.Name).ToArray();
            return Json(domains, JsonRequestBehavior.AllowGet);
        }

        public string DeleteAppServer(int id)
        {
            AppServer server = db.AppServers.Find(id);
            string log = server.Application.Name + " " + server.Name;
            DbSet<AppServer_Report> serverReports = db.AppServerReports;
            foreach (AppServer_Report serverReport in serverReports)
            {
                if (serverReport.AppServerId == server.Id)
                {
                    AppReport report = (AppReport)serverReport.Application_Report.AppReport;
                    DeleteAppReport(report.Id);
                }
            }
            db.AppServers.Remove(server);
            db.SaveChanges();
            Specific_Logging(new Exception(""), "DeleteAppServer", 3);
            return "Le serveur " + server.Name + " a été supprimé";
        }

        [HttpPost]
        public string AddAppServer(int id)
        {
            Application application = db.Applications.Find(id);
            if (application == null)
            {
                return HttpNotFound().ToString();
            }
            string Name = "", Lines = "";
            try
            {
                Name = Request.Form["Name"].ToString();
                Lines = Request.Form["Procedures"].ToString();
                string[] Procedures = Lines.Split(';');
                string[] ReverseProcedures = Procedures.Reverse().ToArray();
                if (db.AppServers.Where(ser => ser.Name == Name.ToUpper() && ser.ApplicationId == application.Id).Count() != 0)
                {
                    return "KO : Le serveur " + Name.ToUpper() + " existe déjà pour cette application, modifiez le directement SVP.";
                }
                Dictionary<int, ServersController.VirtualizedServer> FOREST = ServersController.GetInformationsFromForestDomains();
                ReftechServers[] REFTECH_SERVERS = null;
                try
                {
                    REFTECH_SERVERS = db.ReftechServers.ToArray();
                }
                catch { }
                AppServer server = db.AppServers.Create();
                server.Name = Name.ToUpper();
                ServersController.VirtualizedServer_Result virtual_server = ServersController.GetServerInformations(FOREST, REFTECH_SERVERS, server, HomeController.APP_MODULE, false);
                server.StartOrder = server.StopOrder = "";
                server.ApplicationId = application.Id;
                //StartOrder
                foreach (string Procedure in Procedures)
                {
                    if (Procedure.Trim() != "")
                    {
                        string[] Details = Procedure.Split('|');
                        string type = Details[0].Trim();
                        string target = Details[1].Trim();
                        string action = Details[2].Trim();
                        server.StartOrder += type + "|" + target + "|" + action + ";";
                    }
                }
                server.StartOrder = server.StartOrder.Substring(0, server.StartOrder.Length - 1);
                //StopOrder
                foreach (string Procedure in ReverseProcedures)
                {
                    if (Procedure.Trim() != "")
                    {
                        string[] Details = Procedure.Split('|');
                        string type = Details[0].Trim();
                        string target = Details[1].Trim();
                        string action = Details[2].Trim();
                        if (type == "BATCH")
                        {
                            continue;
                        }
                        switch (action)
                        {
                            case "START":
                                server.StopOrder += type + "|" + target + "|STOP;";
                                break;
                            case "RESTART":
                                server.StopOrder += type + "|" + target + "|STOP;";
                                break;
                            case "CHECK":
                                continue;
                        }
                    }
                }
                if (server.StopOrder.Length > 0)
                {
                    server.StopOrder = server.StopOrder.Substring(0, server.StopOrder.Length - 1);
                }
                if (ModelState.IsValid)
                {
                    db.AppServers.Add(server);
                    db.SaveChanges();
                    Specific_Logging(new Exception(""), "AddAppServer " + server.Application.Name, 3);
                    return "OK";
                }
                return "KO";
            }
            catch (Exception exception)
            {
                Specific_Logging(exception, "AddAppServer");
                return "Erreur lors de l'ajout du serveur";
            }
        }

        [HttpPost]
        public string StartEditAppServer(int id)
        {
            try
            {
                AppServer server = db.AppServers.Find(id);
                string Name = Request.Form["Name"].ToString();
                string Lines = Request.Form["Procedures"].ToString();
                server.Name = Name;
                server.StartOrder = "";
                string[] Procedures = Lines.Split(';');
                foreach (string Procedure in Procedures)
                {
                    if (Procedure.Trim() != "")
                    {
                        string[] Details = Procedure.Split('|');
                        string type = Details[0].Trim();
                        string target = Details[1].Trim();
                        string action = Details[2].Trim();
                        server.StartOrder += type + "|" + target + "|" + action + ";";
                    }
                }
                server.StartOrder = server.StartOrder.Substring(0, server.StartOrder.Length - 1);
                if (ModelState.IsValid)
                {
                    db.Entry(server).State = System.Data.Entity.EntityState.Modified;
                    db.SaveChanges();
                    Specific_Logging(new Exception(""), "StartEditAppServer " + server.Application.Name, 3);
                    return "OK";
                }
                return "KO";
            }
            catch (Exception exception)
            {
                Specific_Logging(exception, "StartEditAppServer");
                return "Erreur lors de la modification de l'application";
            }
        }

        [HttpPost]
        public string StopEditAppServer(int id)
        {
            try
            {
                AppServer server = db.AppServers.Find(id);
                string Name = Request.Form["Name"].ToString();
                string Lines = Request.Form["Procedures"].ToString();
                server.Name = Name;
                server.StopOrder = "";
                string[] ReverseProcedures = Lines.Split(';');
                foreach (string Procedure in ReverseProcedures)
                {
                    if (Procedure.Trim() != "")
                    {
                        string[] Details = Procedure.Split('|');
                        string type = Details[0].Trim();
                        string target = Details[1].Trim();
                        string action = Details[2].Trim();
                        server.StopOrder += type + "|" + target + "|" + action + ";";
                    }
                }
                server.StopOrder = server.StopOrder.Substring(0, server.StopOrder.Length - 1);
                if (ModelState.IsValid)
                {
                    db.Entry(server).State = System.Data.Entity.EntityState.Modified;
                    db.SaveChanges();
                    Specific_Logging(new Exception(""), "StopEditAppServer " + server.Application.Name, 3);
                    return "OK";
                }
                return "KO";
            }
            catch (Exception exception)
            {
                Specific_Logging(exception, "StartEditAppServer");
                return "Erreur lors de la modification de l'application";
            }
        }

        public JsonResult GetAppServerState(int id)
        {
            AppServer server = db.AppServers.Find(id);
            Dictionary<string, string> response = new Dictionary<string, string>();
            response.Add("status", "");
            response.Add("details", "");
            if (server == null)
            {
                response["status"] = "KO";
                response["details"] = "Erreur serveur non trouvé";
                return Json(response, JsonRequestBehavior.AllowGet);
            }
            AppServerInfo serverinfo = new AppServerInfo(server, server.Application);
            try
            {
                return Json(serverinfo.GetState(), JsonRequestBehavior.AllowGet);
            }
            catch (Exception exception)
            {
                Specific_Logging(exception, "GetAppServerState");
                response["status"] = "KO";
                response["details"] = "Exception survenue\n" + exception.Message;
                return Json(response, JsonRequestBehavior.AllowGet);
            }
        }

        public JsonResult StartAppServer(int id)
        {
            AppServer server = db.AppServers.Find(id);
            Dictionary<string, string> response = new Dictionary<string, string>();
            response.Add("status", "");
            response.Add("details", "");
            if (server == null)
            {
                response["status"] = "KO";
                response["details"] = "Erreur serveur non trouvé";
                return Json(response, JsonRequestBehavior.AllowGet);
            }
            AppServerInfo serverinfo = new AppServerInfo(server, server.Application);
            try
            {
                return Json(serverinfo.Start(), JsonRequestBehavior.AllowGet);
            }
            catch (Exception exception)
            {
                Specific_Logging(exception, "StartAppServer " + server.Application.Name);
                response["status"] = "KO";
                response["details"] = "Exception survenue\n" + exception.Message;
                return Json(response, JsonRequestBehavior.AllowGet);
            }
        }

        public JsonResult RestartAppServer(int id)
        {
            AppServer server = db.AppServers.Find(id);
            Dictionary<string, string> response = new Dictionary<string, string>();
            response.Add("status", "");
            response.Add("details", "");
            if (server == null)
            {
                response["status"] = "KO";
                response["details"] = "Erreur serveur non trouvé";
                return Json(response, JsonRequestBehavior.AllowGet);
            }
            AppServerInfo serverinfo = new AppServerInfo(server, server.Application);
            try
            {
                return Json(serverinfo.Restart(), JsonRequestBehavior.AllowGet);
            }
            catch (Exception exception)
            {
                Specific_Logging(new Exception(""), "StopAppServer " + server.Application.Name);
                response["status"] = "KO";
                response["details"] = "Exception survenue\n" + exception.Message;
                return Json(response, JsonRequestBehavior.AllowGet);
            }
        }

        public JsonResult StopAppServer(int id)
        {
            AppServer server = db.AppServers.Find(id);
            Dictionary<string, string> response = new Dictionary<string, string>();
            response.Add("status", "");
            response.Add("details", "");
            if (server == null)
            {
                response["status"] = "KO";
                response["details"] = "Erreur serveur non trouvé";
                return Json(response, JsonRequestBehavior.AllowGet);
            }
            AppServerInfo serverinfo = new AppServerInfo(server, server.Application);
            try
            {
                return Json(serverinfo.Stop(), JsonRequestBehavior.AllowGet);
            }
            catch (Exception exception)
            {
                Specific_Logging(exception, "StopAppServer " + server.Application.Name);
                response["status"] = "KO";
                response["details"] = "Exception survenue\n" + exception.Message;
                return Json(response, JsonRequestBehavior.AllowGet);
            }
        }

        public string Purge()
        {
            string message = "";
            List<AppReport> reports = db.AppReports.Where(rep => rep.Duration == null || rep.ResultPath == null).ToList();
            foreach (AppReport report in reports)
            {
                message += "Rapport " + report.DateTime + " supprimé\n";
                Email email = (report.Email != null) ? report.Email : null;
                List<Application_Report> applicationreports = report.Application_Reports.ToList();
                foreach (Application_Report applicationreport in applicationreports)
                {
                    List<AppServer_Report> serverreports = applicationreport.AppServer_Reports.ToList();
                    foreach (AppServer_Report serverreport in serverreports)
                    {
                        db.AppServerReports.Remove(serverreport);
                    }
                    db.ApplicationReports.Remove(applicationreport);
                }
                db.SaveChanges();
                db.AppReports.Remove(report);
                if (email != null)
                {
                    db.Emails.Remove(email);
                }
                db.SaveChanges();
            }
            Specific_Logging(new Exception(""), "Purge", 3);
            return message;
        }

        public string DeleteAppReport(int id)
        {
            try
            {
                AppReport report = db.AppReports.Find(id);
                Email email = report.Email;

                DbSet<AppServer_Report> applicationserverreports = db.AppServerReports;
                DbSet<Application_Report> applicationappreports = db.ApplicationReports;
                foreach (AppServer_Report appserverreport in applicationserverreports)
                {
                    if (appserverreport.Application_Report.AppReport == report)
                    {
                        db.AppServerReports.Remove(appserverreport);
                    }
                }
                foreach (Application_Report applicationappreport in applicationappreports)
                {
                    if (applicationappreport.AppServer_Reports.Count == 0)
                    {
                        db.ApplicationReports.Remove(applicationappreport);
                    }
                }
                db.Emails.Remove(email);
                System.IO.File.Delete(report.ResultPath);
                db.AppReports.Remove(report);
                db.SaveChanges();
                Specific_Logging(new Exception(""), "DeleteAppReport", 3);
                return "Le rapport a été correctement supprimé";
            }
            catch (Exception exception)
            {
                Specific_Logging(exception, "DeleteAppReport");
                return "Une erreur est surveunue lors de la suppression" +
                    exception.Message;
            }
        }

        public JsonResult CheckApplications()
        {
            if (!CanCheck())
            {
               // return NotifyImpossibility(false);
            }
            Dictionary<string, string> results = new Dictionary<string, string>();
            results["response"] = "OK";
            results["email"] = "";
            results["errors"] = "";
            results["applications"] = "";

            string[] applicationList;
            List<Application> SelectedApps = new List<Application>();
            List<AppDomain> domains = new List<AppDomain>();
            try
            {
                results["applications"] = Request.Form["list"].ToString();
                applicationList = results["applications"].Split(';');
                foreach (string appId in applicationList)
                {
                    int id = 0;
                    Int32.TryParse(appId.Split('-')[1], out id);
                    Application application = db.Applications.Find(id);
                    if (application != null)
                    {
                        SelectedApps.Add(application);
                        AppDomain domain = db.AppDomains.Find(application.Domain);
                        if (!domains.Contains(domain))
                        {
                            domains.Add(domain);
                        }
                    }
                }
            }
            catch (Exception exception)
            {
                results["response"] = "Une erreur est survenue lors de la sélection des Applications";
                results["email"] = "";
                results["errors"] = exception.Message;
                return Json(results, JsonRequestBehavior.AllowGet);
            }
            if (SelectedApps.Count == 0)
            {
                results["response"] = "Aucune Application n'a été sélectionnée dans la base de données";
                results["email"] = "";
                results["errors"] = "";
                return Json(results, JsonRequestBehavior.AllowGet);
            }

            int emailId = 0;
            string ExecutionErrors = "";
            try
            {
                MyApplication = new Excel.Application();
                MyApplication.Visible = false;
                MyWorkbook = MyApplication.Workbooks.Add(Excel.XlWBATemplate.xlWBATWorksheet);
                MySheet = (Excel.Worksheet)MyWorkbook.Sheets.Add();
                MySheet.Name = "Check" + DateTime.Now.ToString("dd") +
                    DateTime.Now.ToString("MM") + DateTime.Now.ToString("yyyy");
                MySheet.Activate();
                int a = 3;
                AppReport report = db.AppReports.Create();
                report.DateTime = DateTime.Now;
                report.TotalChecked = 0;
                report.TotalErrors = 0;
                report.ResultPath = "";
                report.Author = User.Identity.Name;
                report.Module = HomeController.APP_MODULE;

                Email email = db.Emails.Create();
                report.Email = email;
                email.Report = report;
                email.Module = HomeController.APP_MODULE;
                email.Recipients = "";
                email = Emails_Controller.SetRecipients(email, HomeController.APP_MODULE);
                if (ModelState.IsValid)
                {
                    db.AppReports.Add(report);
                    db.SaveChanges();
                    emailId = report.Email.Id;
                    int reportNumber = db.AppReports.Count();
                    if (reportNumber > HomeController.APP_MAX_REPORT_NUMBER)
                    {
                        int reportNumberToDelete = reportNumber - HomeController.APP_MAX_REPORT_NUMBER;
                        AppReport[] reportsToDelete =
                            db.AppReports.OrderBy(id => id.Id).Take(reportNumberToDelete).ToArray();
                        foreach (AppReport toDeleteReport in reportsToDelete)
                        {
                            DeleteAppReport(toDeleteReport.Id);
                        }
                    }
                }
                else
                {
                    results["response"] = "KO";
                    results["email"] = null;
                    results["errors"] = "Impossible de créer un rapport dans la base de données.";
                }

                //START OF TREATMENT
                int lastRow = MySheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row;
                SelectedApps = SelectedApps.OrderBy(app => app.Name).ToList();

                foreach (AppDomain domain in domains.OrderBy(dom => dom.Name))
                {
                    List<Application> dom_applications = SelectedApps.Where(app => app.Domain == domain.Id).ToList();
                    MySheet.Cells[lastRow, 1] = domain.Name;
                    int firstline = lastRow;
                    int mergedlines = 0;
                    foreach (Application application in dom_applications.OrderBy(app => app.Name))
                    {
                        Application_Report applicationReport = new Application_Report();
                        applicationReport.Application = application;
                        applicationReport.AppReport = report;
                        applicationReport.State = "KO";
                        applicationReport.Details = "";
                        applicationReport.Authentified = "";
                        applicationReport.Linkable = "";
                        applicationReport.AppReportId = report.Id;
                        if (ModelState.IsValid)
                        {
                            db.ApplicationReports.Add(applicationReport);
                            db.SaveChanges();
                        }
                        report.TotalChecked++;
                        if (applicationReport.Authentified.Trim() == "")
                        {
                            List<Application> thisApplication = new List<Application>();
                            thisApplication.Add(application);

                            Dictionary<Application, BrowseUrlResult> BrowseResults
                                = new Dictionary<Application, BrowseUrlResult>();
                            BrowseResults = BrowseApplications(thisApplication);

                            applicationReport.Authentified = (BrowseResults[application].Status == "OK") ? "OK" : "KO : " + BrowseResults[application].Details;

                        }
                        if (applicationReport.Authentified != "")
                        {
                            MySheet.Cells[lastRow, 6] = "Authentification : " + applicationReport.Authentified;
                        }
                        else
                        {
                            MySheet.Cells[lastRow, 6] = "Authentification : Inconnue";
                        }

                        AppServer[] servers = application.AppServers.ToArray();
                        Excel.Range ActualRange = MySheet.get_Range("A" + lastRow,
                                "Z" + lastRow);

                        int lines = servers.Length;
                        lines = (lines == 0) ? 1 : lines;
                        mergedlines += lines;

                        if (servers.Length == 0)
                        {
                            MySheet.Cells[lastRow, 2] = application.Name;
                            MySheet.Cells[lastRow, 2].EntireColumn.ColumnWidth = 25;
                            MySheet.Cells[lastRow, 2].EntireRow.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                            MySheet.Cells[lastRow, 2].EntireRow.Font.Color = System.Drawing.ColorTranslator.FromHtml("#000");
                            MySheet.Cells[lastRow, 2].EntireRow.Style.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                            MySheet.Cells[lastRow, 2].EntireRow.Style.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                            MySheet.Cells[lastRow, 2].EntireRow.Font.Bold = true;
                            MySheet.Cells[lastRow, 2].EntireRow.Interior.Color = System.Drawing.ColorTranslator.FromHtml("#eeece1");
                            MySheet.Hyperlinks.Add(MySheet.Cells[lastRow, 2], application.Url, Type.Missing, application.Name, application.Name);

                            MySheet.Cells[lastRow, 6].Interior.Color = System.Drawing.ColorTranslator.FromHtml("#fff");
                            MySheet.Cells[lastRow, 6].EntireColumn.ColumnWidth = 40;
                            MySheet.Cells[lastRow, 6].Font.Color = (applicationReport.Authentified == "OK") ? System.Drawing.ColorTranslator.FromHtml("#22b14c")
                                   : System.Drawing.ColorTranslator.FromHtml("#ff2f00");

                            MySheet.Cells[lastRow, 3].EntireColumn.ColumnWidth = 15;
                            MySheet.Cells[lastRow, 4].EntireColumn.ColumnWidth = 10;
                            MySheet.Cells[lastRow, 5].EntireColumn.ColumnWidth = 10;
                            MySheet.Cells[lastRow, 5].Font.Color = System.Drawing.ColorTranslator.FromHtml("#de5a26");

                            applicationReport.State = "";


                            if (applicationReport.Authentified != "")
                            {
                                MySheet.Cells[lastRow, 6] = "Authentification : " + applicationReport.Authentified;
                            }
                            else
                            {
                                MySheet.Cells[lastRow, 6] = "Authentification : Inconnue";
                            }
                            if (ModelState.IsValid)
                            {
                                db.Entry(applicationReport).State = System.Data.Entity.EntityState.Modified;
                                db.SaveChanges();
                            }
                            lastRow += 1;
                        }
                        else
                        {
                            MySheet.Cells[lastRow, 2] = application.Name;
                            MySheet.Cells[lastRow, 2].EntireColumn.ColumnWidth = 25;
                            MySheet.Cells[lastRow, 6].Font.Color = (applicationReport.Authentified == "OK") ? System.Drawing.ColorTranslator.FromHtml("#22b14c")
                                   : System.Drawing.ColorTranslator.FromHtml("#ff2f00");

                            Excel.Range to_merge = MySheet.get_Range("B" + lastRow, "B" + (lastRow + application.AppServers.Count - 1));
                            to_merge.Merge();
                            to_merge.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                            to_merge.Font.Bold = true;
                            to_merge.Style.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                            to_merge.Style.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

                            Excel.Range auth_to_merge = MySheet.get_Range("F" + lastRow, "F" + (lastRow + application.AppServers.Count - 1));
                            auth_to_merge.Merge();
                            auth_to_merge.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                            auth_to_merge.Font.Bold = true;
                            auth_to_merge.Style.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                            auth_to_merge.Style.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

                            List<string> Status = new List<string>();
                            foreach (AppServer server in servers)
                            {
                                MySheet.Cells[lastRow, 3] = server.Name;
                                MySheet.Cells[lastRow, 3].EntireRow.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                                MySheet.Cells[lastRow, 3].EntireRow.Font.Color = System.Drawing.ColorTranslator.FromHtml("#000");
                                MySheet.Cells[lastRow, 3].EntireRow.Style.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                                MySheet.Cells[lastRow, 3].EntireRow.Style.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                                MySheet.Cells[lastRow, 3].EntireRow.Font.Bold = true;

                                MySheet.Cells[lastRow, 3].EntireColumn.ColumnWidth = 15;



                                JsonResult JsonResponse = GetAppServerState(server.Id);
                                Dictionary<string, string> response = (Dictionary<string, string>)JsonResponse.Data;

                                AppServer_Report serverreport = db.AppServerReports.Create();
                                serverreport.AppServer = server;
                                serverreport.Application_Report = applicationReport;
                                serverreport.State = serverreport.Details = serverreport.Ping = "";// response["status"];
                                serverreport.State = response["status"];


                                if (serverreport.State != "OK")
                                {
                                    Status.Add("KO");
                                    try
                                    {
                                        Ping ping = new Ping();
                                        PingOptions options = new PingOptions(64, true);
                                        PingReply pingreply = ping.Send(server.Name);
                                        serverreport.Ping = (pingreply.Status.ToString() == "Success") ? "OK" : "KO";
                                    }
                                    catch
                                    {
                                        serverreport.Ping = "KO";
                                    }
                                    MySheet.Cells[lastRow, 3].EntireRow.Interior.Color = System.Drawing.ColorTranslator.FromHtml("#ff3f3f");
                                    MySheet.Cells[lastRow, 5].Interior.Color = System.Drawing.ColorTranslator.FromHtml("#fff");
                                    MySheet.Cells[lastRow, 5].EntireColumn.ColumnWidth = 10;
                                    MySheet.Cells[lastRow, 4] = "KO";
                                    MySheet.Cells[lastRow, 5] = "Ping: " + serverreport.Ping;

                                    MySheet.Cells[lastRow, 5].Font.Color = (serverreport.Ping == "OK") ? System.Drawing.ColorTranslator.FromHtml("#22b14c")
                                        : System.Drawing.ColorTranslator.FromHtml("#ff2f00");
                                    serverreport.Details = "";
                                    int column = 7;
                                    string[] Details = response["details"].Split(new string[] { "\n" }, StringSplitOptions.None);
                                    foreach (string infos in Details)
                                    {
                                        if (infos.Trim() != "")
                                        {
                                            serverreport.Details += infos + " | ";
                                            MySheet.Cells[lastRow, column] = infos;
                                            MySheet.Cells[lastRow, column].EntireColumn.ColumnWidth = 35;
                                            column++;
                                        }
                                    }
                                    if (serverreport.Details.Length > 1)
                                    {
                                        serverreport.Details = serverreport.Details.Substring(0, serverreport.Details.Length - 3);
                                    }
                                }
                                else
                                {
                                    Status.Add("OK");
                                    serverreport.Ping = "OK";
                                    MySheet.Cells[lastRow, 3].EntireRow.Interior.Color = System.Drawing.ColorTranslator.FromHtml("#22b14c");
                                    MySheet.Cells[lastRow, 5].Interior.Color = System.Drawing.ColorTranslator.FromHtml("#fff");
                                    MySheet.Cells[lastRow, 5].EntireColumn.ColumnWidth = 10;
                                    MySheet.Cells[lastRow, 4] = "OK";
                                    MySheet.Cells[lastRow, 5] = "Ping: OK";
                                    MySheet.Cells[lastRow, 5].Font.Color = System.Drawing.ColorTranslator.FromHtml("#22b14c");
                                    serverreport.Details = "";
                                }

                                if (ModelState.IsValid)
                                {
                                    db.AppServerReports.Add(serverreport);
                                    db.SaveChanges();
                                }

                                auth_to_merge.Style.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                                auth_to_merge.Interior.Color = System.Drawing.ColorTranslator.FromHtml("#fff");
                                auth_to_merge.EntireColumn.ColumnWidth = 40;

                                auth_to_merge.Font.Color = (applicationReport.Authentified == "OK") ? System.Drawing.ColorTranslator.FromHtml("#22b14c")
                                           : System.Drawing.ColorTranslator.FromHtml("#ff2f00");
                                auth_to_merge.EntireColumn.ColumnWidth = 40;
                                if (applicationReport.Authentified != "")
                                {
                                    MySheet.Cells[lastRow, 6] = "Authentification : " + applicationReport.Authentified;
                                }
                                else
                                {
                                    MySheet.Cells[lastRow, 6] = "Authentification : Inconnue";

                                }
                                lastRow += 1;
                            }
                            to_merge.Hyperlinks.Add(to_merge, application.Url, Type.Missing, application.Name, application.Name);
                            if (!Status.Contains("KO"))
                            {
                                applicationReport.State = "OK";
                            }
                            else
                            {
                                report.TotalErrors++;
                                if (!Status.Contains("OK"))
                                {
                                    applicationReport.State = "KO";
                                }
                                else
                                {
                                    applicationReport.State = "H-OK";
                                }
                            }
                            if (ModelState.IsValid)
                            {
                                db.Entry(applicationReport).State = System.Data.Entity.EntityState.Modified;
                                db.SaveChanges();
                            }
                            string backgroundcolor = (applicationReport.State == "OK") ? "#22b14c" : (applicationReport.State == "KO") ? "#ff3f3f" :
                                (applicationReport.State == "H-OK") ? "#de5a26" : "#eeece1";
                            to_merge.Interior.Color = System.Drawing.ColorTranslator.FromHtml(backgroundcolor);
                        }
                    }
                    mergedlines = (mergedlines == 0) ? 1 : mergedlines;
                    Excel.Range dom_to_merge = MySheet.get_Range("A" + firstline, "A" + (firstline + mergedlines - 1));
                    dom_to_merge.Merge();
                    dom_to_merge.Interior.Color = System.Drawing.ColorTranslator.FromHtml("#dcdbdb");
                    dom_to_merge.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                    dom_to_merge.Font.Bold = true;
                    dom_to_merge.EntireColumn.ColumnWidth = 35;
                    dom_to_merge.Style.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                    dom_to_merge.Style.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                }
                string FileName = HomeController.APP_RESULTS_FOLDER + "Check Applications " + DateTime.Now.ToString("dd") +
                    DateTime.Now.ToString("MM") + DateTime.Now.ToString("yyyy") + " - " + report.Id + ".xlsx";
                MyWorkbook.SaveAs(FileName,
                    Excel.XlFileFormat.xlOpenXMLWorkbook, null, null, null, null,
                    Excel.XlSaveAsAccessMode.xlNoChange, null, null, null, null, null);
                report.Duration = DateTime.Now.Subtract(report.DateTime);
                report.ResultPath = FileName;
                if (ModelState.IsValid)
                {
                    db.Entry(report).State = System.Data.Entity.EntityState.Modified;
                    db.SaveChanges();
                    string buildOk = BuildEmail(email.Id);
                    if (buildOk != "BuildOK")
                    {
                        ExecutionErrors += "Erreur lors de la mise à jour du mail \n <br />";
                    }
                }
                else
                {
                    results["response"] = "KO";
                    results["email"] = null;
                    results["errors"] = "Echec lors de l'enregistrement dans la base de données.";
                }
            }
            catch (Exception exception)
            {
                Specific_Logging(exception, "CheckApplications");
            }
            finally
            {
                McoUtilities.CloseExcel(MyApplication, MyWorkbook, MySheet);
            }

            results["email"] = emailId.ToString();
            results["errors"] = "Fin d'exécution. \n" + "Erreurs d'exécution : " + ExecutionErrors;
            Specific_Logging(new Exception(""), "CheckApplications", 3);
            return Json(results, JsonRequestBehavior.AllowGet);
        }

        public JsonResult FurtherCheckApplications()
        {
            if (!CanCheck())
            {
                return NotifyImpossibility(false);
            }
            Dictionary<string, string> results = new Dictionary<string, string>();
            results["response"] = "OK";
            results["email"] = "";
            results["errors"] = "";
            results["applications"] = "";
            results["servers"] = "";

            string[] applicationList, serverList;
            List<Application> SelectedApps = new List<Application>();
            List<AppServer> SelectedServers = new List<AppServer>();
            List<AppDomain> domains = new List<AppDomain>();
            try
            {
                results["applications"] = Request.Form["list"].ToString();
                applicationList = results["applications"].Split(';');
                results["servers"] = Request.Form["server_list"].ToString();
                serverList = results["servers"].Split(';');
                foreach (string appId in applicationList)
                {
                    int id = 0;
                    Int32.TryParse(appId.Split('-')[1], out id);
                    Application application = db.Applications.Find(id);
                    if (application != null)
                    {
                        SelectedApps.Add(application);
                        AppDomain domain = db.AppDomains.Find(application.Domain);
                        if (!domains.Contains(domain))
                        {
                            domains.Add(domain);
                        }
                    }
                }
                foreach (string serverId in serverList)
                {
                    int id = 0;
                    Int32.TryParse(serverId, out id);
                    AppServer server = db.AppServers.Find(id);
                    if (server != null)
                    {
                        SelectedServers.Add(server);
                    }
                }
            }
            catch (Exception exception)
            {
                results["response"] = "Une erreur est survenue lors de la sélection des Applications";
                results["email"] = "";
                results["errors"] = exception.Message;
                return Json(results, JsonRequestBehavior.AllowGet);
            }
            if (SelectedApps.Count == 0)
            {
                results["response"] = "Aucune Application n'a été sélectionnée dans la base de données";
                results["email"] = "";
                results["errors"] = "";
                return Json(results, JsonRequestBehavior.AllowGet);
            }

            int emailId = 0;
            string ExecutionErrors = "";
            try
            {
                MyApplication = new Excel.Application();
                MyApplication.Visible = false;
                MyWorkbook = MyApplication.Workbooks.Add(Excel.XlWBATemplate.xlWBATWorksheet);
                MySheet = (Excel.Worksheet)MyWorkbook.Sheets.Add();
                MySheet.Name = "Check" + DateTime.Now.ToString("dd") +
                    DateTime.Now.ToString("MM") + DateTime.Now.ToString("yyyy");
                MySheet.Activate();

                AppReport report = db.AppReports.Create();
                report.DateTime = DateTime.Now;
                report.TotalChecked = 0;
                report.TotalErrors = 0;
                report.ResultPath = "";
                report.Module = HomeController.APP_MODULE;
                report.Author = User.Identity.Name;

                Email email = db.Emails.Create();
                report.Email = email;
                email.Report = report;
                email.Recipients = "";
                email.Module = HomeController.APP_MODULE;
                email = Emails_Controller.SetRecipients(email, HomeController.APP_MODULE);
                if (ModelState.IsValid)
                {
                    db.AppReports.Add(report);
                    db.SaveChanges();
                    emailId = report.Email.Id;
                    int reportNumber = db.AppReports.Count();
                    if (reportNumber > HomeController.APP_MAX_REPORT_NUMBER)
                    {
                        int reportNumberToDelete = reportNumber - HomeController.APP_MAX_REPORT_NUMBER;
                        AppReport[] reportsToDelete =
                            db.AppReports.OrderBy(id => id.Id).Take(reportNumberToDelete).ToArray();
                        foreach (AppReport toDeleteReport in reportsToDelete)
                        {
                            DeleteAppReport(toDeleteReport.Id);
                        }
                    }
                }
                else
                {
                    results["response"] = "KO";
                    results["email"] = null;
                    results["errors"] = "Impossible de créer un rapport dans la base de données.";
                }

                //START OF TREATMENT
                int lastRow = MySheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row;
                SelectedApps = SelectedApps.OrderBy(app => app.Name).ToList();

                foreach (AppDomain domain in domains.OrderBy(dom => dom.Name))
                {
                    List<Application> dom_applications = SelectedApps.Where(app => app.Domain == domain.Id).ToList();
                    MySheet.Cells[lastRow, 1] = domain.Name;
                    int firstline = lastRow;
                    int mergedlines = 0;
                    foreach (Application application in dom_applications.OrderBy(app => app.Name))
                    {
                        Application_Report applicationReport = new Application_Report();
                        applicationReport.Application = application;
                        applicationReport.AppReport = report;
                        applicationReport.State = "KO";
                        applicationReport.Details = "";
                        applicationReport.Authentified = "";
                        applicationReport.Linkable = "";
                        applicationReport.AppReportId = report.Id;
                        if (ModelState.IsValid)
                        {
                            db.ApplicationReports.Add(applicationReport);
                            db.SaveChanges();
                        }
                        report.TotalChecked++;
                        if (applicationReport.Authentified.Trim() == "")
                        {
                            List<Application> thisApplication = new List<Application>();
                            thisApplication.Add(application);

                            Dictionary<Application, BrowseUrlResult> BrowseResults
                                = new Dictionary<Application, BrowseUrlResult>();
                            BrowseResults = BrowseApplications(thisApplication);

                            applicationReport.Authentified = (BrowseResults[application].Status == "OK") ? "OK" : "KO : " + BrowseResults[application].Details;
                        }

                        AppServer[] servers = application.AppServers.ToArray();
                        Excel.Range ActualRange = MySheet.get_Range("A" + lastRow,
                                "Z" + lastRow);

                        int lines = servers.Length;
                        lines = (lines == 0) ? 1 : lines;
                        mergedlines += lines;

                        if (servers.Length == 0)
                        {
                            MySheet.Cells[lastRow, 2] = application.Name;
                            MySheet.Cells[lastRow, 2].EntireColumn.ColumnWidth = 25;
                            MySheet.Cells[lastRow, 2].EntireRow.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                            MySheet.Cells[lastRow, 2].EntireRow.Font.Color = System.Drawing.ColorTranslator.FromHtml("#000");
                            MySheet.Cells[lastRow, 2].EntireRow.Style.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                            MySheet.Cells[lastRow, 2].EntireRow.Style.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                            MySheet.Cells[lastRow, 2].EntireRow.Font.Bold = true;
                            MySheet.Cells[lastRow, 2].EntireRow.Interior.Color = System.Drawing.ColorTranslator.FromHtml("#eeece1");
                            MySheet.Hyperlinks.Add(MySheet.Cells[lastRow, 2], application.Url, Type.Missing, application.Name, application.Name);

                            MySheet.Cells[lastRow, 6].Interior.Color = System.Drawing.ColorTranslator.FromHtml("#fff");
                            MySheet.Cells[lastRow, 6].EntireColumn.ColumnWidth = 40;

                            MySheet.Cells[lastRow, 3].EntireColumn.ColumnWidth = 15;
                            MySheet.Cells[lastRow, 4].EntireColumn.ColumnWidth = 10;
                            MySheet.Cells[lastRow, 5].EntireColumn.ColumnWidth = 10;
                            MySheet.Cells[lastRow, 5].Font.Color = System.Drawing.ColorTranslator.FromHtml("#de5a26");

                            applicationReport.State = "";

                            if (applicationReport.Authentified != "")
                            {
                                MySheet.Cells[lastRow, 6] = "Authentification : " + applicationReport.Authentified;
                            }
                            else
                            {
                                MySheet.Cells[lastRow, 6] = "Authentification : Inconnue";
                            }
                            if (ModelState.IsValid)
                            {
                                db.Entry(applicationReport).State = System.Data.Entity.EntityState.Modified;
                                db.SaveChanges();
                            }
                            lastRow += 1;
                        }
                        else
                        {
                            MySheet.Cells[lastRow, 2] = application.Name;

                            Excel.Range to_merge = MySheet.get_Range("B" + lastRow, "B" + (lastRow + application.AppServers.Count - 1));
                            to_merge.Merge();
                            to_merge.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                            to_merge.Font.Bold = true;
                            to_merge.Style.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                            to_merge.Style.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

                            Excel.Range auth_to_merge = MySheet.get_Range("F" + lastRow, "F" + (lastRow + application.AppServers.Count - 1));
                            auth_to_merge.Merge();
                            auth_to_merge.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                            auth_to_merge.Font.Bold = true;
                            auth_to_merge.Style.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                            auth_to_merge.Style.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

                            List<string> Status = new List<string>();
                            foreach (AppServer server in servers)
                            {
                                MySheet.Cells[lastRow, 2].EntireColumn.ColumnWidth = 25;
                                MySheet.Cells[lastRow, 2].EntireRow.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                                MySheet.Cells[lastRow, 2].EntireRow.Font.Color = System.Drawing.ColorTranslator.FromHtml("#000");
                                MySheet.Cells[lastRow, 2].EntireRow.Style.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                                MySheet.Cells[lastRow, 2].EntireRow.Style.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                                MySheet.Cells[lastRow, 2].EntireRow.Font.Bold = true;

                                MySheet.Cells[lastRow, 3] = server.Name;
                                MySheet.Cells[lastRow, 3].EntireColumn.ColumnWidth = 15;


                                JsonResult JsonResponse = GetAppServerState(server.Id);
                                Dictionary<string, string> response = (Dictionary<string, string>)JsonResponse.Data;
                                AppServer_Report serverreport = db.AppServerReports.Create();
                                serverreport.AppServer = server;
                                serverreport.Application_Report = applicationReport;
                                serverreport.State = response["status"];


                                //CHECK URL
                                List<string> serverAuths = new List<string>();
                                if (SelectedServers.Contains(server))
                                {
                                    List<AppServer> thisServer = new List<AppServer>();
                                    thisServer.Add(server);
                                    Dictionary<AppServer, List<BrowseUrlResult>> BrowseServerResults
                                        = new Dictionary<AppServer, List<BrowseUrlResult>>();
                                    BrowseServerResults = BrowseServers(thisServer);
                                    foreach (BrowseUrlResult checkedUrl in BrowseServerResults[server])
                                    {
                                        string result = checkedUrl.Target + " --> ";
                                        result += (checkedUrl.Status == "OK") ? "OK" : "KO : " + checkedUrl.Details;
                                        serverAuths.Add(result);
                                    }
                                }
                                //END CHECK URL
                                int last_error_row = 7;
                                if (serverreport.State != "OK")
                                {
                                    Status.Add("KO");
                                    try
                                    {
                                        Ping ping = new Ping();
                                        PingOptions options = new PingOptions(64, true);
                                        PingReply pingreply = ping.Send(server.Name);
                                        serverreport.Ping = (pingreply.Status.ToString() == "Success") ? "OK" : "KO";
                                    }
                                    catch
                                    {
                                        serverreport.Ping = "KO";
                                    }
                                    MySheet.Cells[lastRow, 2].EntireRow.Interior.Color = System.Drawing.ColorTranslator.FromHtml("#ff3f3f");
                                    MySheet.Cells[lastRow, 5].Interior.Color = System.Drawing.ColorTranslator.FromHtml("#fff");
                                    MySheet.Cells[lastRow, 5].EntireColumn.ColumnWidth = 10;
                                    MySheet.Cells[lastRow, 4] = "KO";
                                    MySheet.Cells[lastRow, 5] = "Ping: " + serverreport.Ping;

                                    MySheet.Cells[lastRow, 5].Font.Color = (serverreport.Ping == "OK") ? System.Drawing.ColorTranslator.FromHtml("#22b14c")
                                        : System.Drawing.ColorTranslator.FromHtml("#ff2f00");
                                    serverreport.Details = "";
                                    int column = 7;
                                    string[] Details = response["details"].Split(new string[] { "\n" }, StringSplitOptions.None);
                                    foreach (string infos in Details)
                                    {
                                        if (infos.Trim() != "")
                                        {
                                            serverreport.Details += infos + " | ";
                                            MySheet.Cells[lastRow, column] = infos;
                                            MySheet.Cells[lastRow, column].EntireColumn.ColumnWidth = 35;
                                            column++;
                                        }
                                    }
                                    last_error_row = column;
                                    if (serverreport.Details.Length > 3)
                                    {
                                        serverreport.Details = serverreport.Details.Substring(0, serverreport.Details.Length - 3);
                                    }
                                }
                                else
                                {
                                    Status.Add("OK");
                                    serverreport.Ping = "OK";
                                    MySheet.Cells[lastRow, 2].EntireRow.Interior.Color = System.Drawing.ColorTranslator.FromHtml("#22b14c");
                                    MySheet.Cells[lastRow, 5].Interior.Color = System.Drawing.ColorTranslator.FromHtml("#fff");
                                    MySheet.Cells[lastRow, 5].EntireColumn.ColumnWidth = 10;
                                    MySheet.Cells[lastRow, 4] = "OK";
                                    MySheet.Cells[lastRow, 5] = "Ping: OK";
                                    MySheet.Cells[lastRow, 5].Font.Color = System.Drawing.ColorTranslator.FromHtml("#22b14c");
                                    serverreport.Details = "";
                                }
                                if (SelectedServers.Contains(server))
                                {
                                    foreach (string serverAuth in serverAuths)
                                    {
                                        serverreport.Details += (serverreport.Details.Count() == 0) ?
                                            serverAuth : "<br />" + serverAuth;
                                        MySheet.Cells[lastRow, last_error_row] = serverAuth;
                                        MySheet.Cells[lastRow, last_error_row].EntireColumn.ColumnWidth = 40;
                                        last_error_row++;
                                    }
                                }
                                if (ModelState.IsValid)
                                {
                                    db.AppServerReports.Add(serverreport);
                                    db.SaveChanges();
                                }

                                auth_to_merge.Style.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                                auth_to_merge.Interior.Color = System.Drawing.ColorTranslator.FromHtml("#fff");
                                auth_to_merge.EntireColumn.ColumnWidth = 40;

                                auth_to_merge.Font.Color = (applicationReport.Authentified == "OK") ? System.Drawing.ColorTranslator.FromHtml("#22b14c")
                                           : System.Drawing.ColorTranslator.FromHtml("#ff2f00");
                                auth_to_merge.EntireColumn.ColumnWidth = 40;
                                if (applicationReport.Authentified != "")
                                {
                                    MySheet.Cells[lastRow, 6] = "Authentification : " + applicationReport.Authentified;
                                }
                                else
                                {
                                    MySheet.Cells[lastRow, 6] = "Authentification : Inconnue";

                                }
                                lastRow += 1;
                            }
                            if (!Status.Contains("KO"))
                            {
                                applicationReport.State = "OK";
                            }
                            else
                            {
                                report.TotalErrors++;
                                if (!Status.Contains("OK"))
                                {
                                    applicationReport.State = "KO";
                                }
                                else
                                {
                                    applicationReport.State = "H-OK";
                                }
                            }
                            if (ModelState.IsValid)
                            {
                                db.Entry(applicationReport).State = System.Data.Entity.EntityState.Modified;
                                db.SaveChanges();
                            }
                            string backgroundcolor = (applicationReport.State == "OK") ? "#22b14c" : (applicationReport.State == "KO") ? "#ff3f3f" :
                                (applicationReport.State == "H-OK") ? "#de5a26" : "#eeece1";
                            to_merge.Interior.Color = System.Drawing.ColorTranslator.FromHtml(backgroundcolor);
                        }
                    }
                    mergedlines = (mergedlines == 0) ? 1 : mergedlines;
                    Excel.Range dom_to_merge = MySheet.get_Range("A" + firstline, "A" + (firstline + mergedlines - 1));
                    dom_to_merge.Merge();
                    dom_to_merge.Interior.Color = System.Drawing.ColorTranslator.FromHtml("#dcdbdb");
                    dom_to_merge.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                    dom_to_merge.Font.Bold = true;
                    dom_to_merge.EntireColumn.ColumnWidth = 35;
                    dom_to_merge.Style.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                    dom_to_merge.Style.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                }
                string FileName = HomeController.APP_RESULTS_FOLDER + "Check Applications " + DateTime.Now.ToString("dd") +
                    DateTime.Now.ToString("MM") + DateTime.Now.ToString("yyyy") + " - " + report.Id + ".xlsx";
                MyWorkbook.SaveAs(FileName,
                    Excel.XlFileFormat.xlOpenXMLWorkbook, null, null, null, null,
                    Excel.XlSaveAsAccessMode.xlNoChange, null, null, null, null, null);
                report.Duration = DateTime.Now.Subtract(report.DateTime);
                report.ResultPath = FileName;
                if (ModelState.IsValid)
                {
                    db.Entry(report).State = System.Data.Entity.EntityState.Modified;
                    db.SaveChanges();
                    string buildOk = BuildEmail(email.Id);
                    if (buildOk != "BuildOK")
                    {
                        ExecutionErrors += "Erreur lors de la mise à jour du mail \n <br />";
                    }
                }
                else
                {
                    results["response"] = "KO";
                    results["email"] = null;
                    results["errors"] = "Echec lors de l'enregistrement dans la base de données.";
                }
            }
            catch (Exception exception)
            {
                Specific_Logging(exception, "FurtherCheckApplications");
            }
            finally
            {
                McoUtilities.CloseExcel(MyApplication, MyWorkbook, MySheet);
            }
            results["email"] = emailId.ToString();
            results["errors"] = "Fin d'exécution. \n" + "Erreurs d'exécution : " + ExecutionErrors;
            Specific_Logging(new Exception(""), "FurtherCheckApplications", 3);
            return Json(results, JsonRequestBehavior.AllowGet);
        }

        public class Procedure
        {
            public static List<string> TYPES = new List<string>(new string[] { "SERVICE", "PROCESS", "BATCH", "URL" });

            public static List<string> ACTIONS = new List<string>(new string[] { "START", "STOP", "RESTART", "CHECK" });

            public static List<string> STATES = new List<string>(new string[] { "Running", "Stopped", "Found", "Unknown" });

            public string Target { get; set; }

            public string Type { get; set; }

            public string State { get; set; }

            public string Action { get; set; }

            public List<Procedure> Dependencies { get; set; }

            public Procedure(string Target, string Type, string Action, List<Procedure> Dependencies = null)
            {
                if (Procedure.TYPES.Contains(Type) && Procedure.ACTIONS.Contains(Action))
                {
                    this.Target = Target;
                    this.Type = Type;
                    this.Action = Action;
                    if (this.Type == "BATCH")
                    {
                        this.Action = "START";
                    }
                    if (this.Type == "URL")
                    {
                        this.Action = "CHECK";
                    }
                    this.Dependencies = Dependencies;
                    this.State = "Unknown";
                }
            }

            //-------------------------------------------------------------------
            //SERVICE FUNCTIONS
            public bool ServiceStart(string Server)
            {
                IntPtr userToken = IntPtr.Zero;
                bool success = McoUtilities.LogonUser(
                  HomeController.DEFAULT_USERNAME_IMPERSONNATION,
                  HomeController.DEFAULT_DOMAIN_IMPERSONNATION,
                  McoUtilities.Decrypt(HomeController.DEFAULT_PASSWORD_IMPERSONNATION),
                  (int)McoUtilities.LogonType.LOGON32_LOGON_INTERACTIVE, //2
                  (int)McoUtilities.LogonProvider.LOGON32_PROVIDER_DEFAULT, //0
                  out userToken);
                if (!success)
                {
                    return false;
                }

                using (WindowsIdentity.Impersonate(userToken))
                {
                    ServiceController[] services = ServiceController.GetServices(Server);
                    foreach (ServiceController service in services)
                    {
                        try
                        {
                            if (service.DisplayName.ToLower() == this.Target.ToLower() || service.ServiceName.ToLower() == this.Target.ToLower())
                            {
                                ServiceSpecialController targetedService = new ServiceSpecialController(service.ServiceName, Server);
                                if (targetedService.Status != ServiceControllerStatus.Running)
                                {
                                    targetedService.StartupType = "Manual";
                                    targetedService.Start();
                                    targetedService.WaitForStatus(ServiceControllerStatus.Running);
                                }
                                return true;
                            }
                        }
                        catch (Exception exception)
                        {
                            string log_path_info = exception.Message + "\r\n";
                            try
                            {
                                string log = "\r\n**************************************************\r\n";
                                log += DateTime.Now.ToString() + " : " + "Service Start App Error : General Error \r\n";
                                log += exception.Message + "\r\n";
                                System.IO.File.AppendAllText(HomeController.GENERAL_LOG_FILE, log);
                            }
                            catch { }
                            return false;
                        }
                    }
                }
                return false;
            }

            public bool ServiceStop(string Server)
            {
                IntPtr userToken = IntPtr.Zero;
                bool success = McoUtilities.LogonUser(
                  HomeController.DEFAULT_USERNAME_IMPERSONNATION,
                  HomeController.DEFAULT_DOMAIN_IMPERSONNATION,
                  McoUtilities.Decrypt(HomeController.DEFAULT_PASSWORD_IMPERSONNATION),
                  (int)McoUtilities.LogonType.LOGON32_LOGON_INTERACTIVE, //2
                  (int)McoUtilities.LogonProvider.LOGON32_PROVIDER_DEFAULT, //0
                  out userToken);
                if (!success)
                {
                    return false;
                }

                using (WindowsIdentity.Impersonate(userToken))
                {
                    ServiceController[] services = ServiceController.GetServices(Server);
                    foreach (ServiceController service in services)
                    {
                        try
                        {
                            if (service.DisplayName.ToLower() == this.Target.ToLower() || service.ServiceName.ToLower() == this.Target.ToLower())
                            {
                                ServiceSpecialController targetedService = new ServiceSpecialController(service.ServiceName, Server);
                                if (targetedService.Status != ServiceControllerStatus.Stopped)
                                {
                                    targetedService.StartupType = "Manual";
                                    targetedService.Stop();
                                    targetedService.WaitForStatus(ServiceControllerStatus.Stopped);
                                }
                                return true;
                            }
                        }
                        catch (Exception exception)
                        {
                            string log_path_info = exception.Message + "\r\n";
                            try
                            {
                                string log = "\r\n**************************************************\r\n";
                                log += DateTime.Now.ToString() + " : " + "Application ServiceStop Error : General Error \r\n";
                                log += exception.Message + "\r\n";
                                System.IO.File.AppendAllText(HomeController.GENERAL_LOG_FILE, log);
                            }
                            catch { }
                            return false;
                        }
                    }
                }
                return false;
            }

            public bool ServiceRestart(string Server)
            {
                if (this.Stop(Server))
                {
                    return this.Start(Server);
                }
                return false;
            }

            public string ServiceGetState(string Server)
            {
                string answer = "Unknown";
                IntPtr userToken = IntPtr.Zero;
                bool success = McoUtilities.LogonUser(
                  HomeController.DEFAULT_USERNAME_IMPERSONNATION,
                  HomeController.DEFAULT_DOMAIN_IMPERSONNATION,
                  McoUtilities.Decrypt(HomeController.DEFAULT_PASSWORD_IMPERSONNATION),
                  (int)McoUtilities.LogonType.LOGON32_LOGON_INTERACTIVE, //2
                  (int)McoUtilities.LogonProvider.LOGON32_PROVIDER_DEFAULT, //0
                  out userToken);
                if (!success)
                {
                    return "Authentication Error";
                }

                using (WindowsIdentity.Impersonate(userToken))
                {
                    ServiceController[] services = ServiceController.GetServices(Server);
                    foreach (ServiceController service in services)
                    {
                        try
                        {
                            if (service.DisplayName.ToLower() == this.Target.ToLower() || service.ServiceName.ToLower() == this.Target.ToLower())
                            {
                                ServiceSpecialController targetedService = new ServiceSpecialController(service.ServiceName, Server);
                                if (targetedService.Status.Equals(ServiceControllerStatus.Running))
                                {
                                    this.State = "Running";
                                    return "Running";
                                }
                                if (targetedService.Status.Equals(ServiceControllerStatus.Stopped))
                                {
                                    this.State = "Stopped";
                                    return "Stopped";
                                }
                            }
                            else
                            {
                                answer = "Not found";
                            }
                        }
                        catch (Exception exception)
                        {
                            this.State = "Exception " + exception.Message;
                            return "Exception " + exception.Message;
                        }
                    }
                }
                return answer;
            }
            //_______________________________________________________________________________

            //BATCH FUNCTIONS
            public bool BatchStart(string Server)
            {
                IntPtr userToken = IntPtr.Zero;
                bool success = McoUtilities.LogonUser(
                  HomeController.DEFAULT_USERNAME_IMPERSONNATION,
                  HomeController.DEFAULT_DOMAIN_IMPERSONNATION,
                  McoUtilities.Decrypt(HomeController.DEFAULT_PASSWORD_IMPERSONNATION),
                  (int)McoUtilities.LogonType.LOGON32_LOGON_INTERACTIVE, //2
                  (int)McoUtilities.LogonProvider.LOGON32_PROVIDER_DEFAULT, //0
                  out userToken);
                if (!success)
                {
                    return false;
                }
                Process process = new Process();
                process.StartInfo.LoadUserProfile = true;
                process.StartInfo.UseShellExecute = false;
                process.StartInfo.Verb = "Runas";
                process.StartInfo.FileName = "cmd.exe";
                try
                {
                    string ServerFolder = @"\\" + Server;
                    string FileName = "";
                    string[] ServerInfos = this.Target.Split(':');
                    if (ServerInfos[0].Length == 1 && ServerInfos.Length == 2)
                    {
                        ServerFolder += "\\" + ServerInfos[0] + "$";
                        string[] Path = ServerInfos[1].Split('\\');
                        for (int index = 0; index < Path.Length - 1; index++)
                        {
                            ServerFolder += "\\" + Path[index];
                        }
                        FileName = Path[Path.Length - 1];
                    }

                    process.StartInfo.Arguments = @"/c powershell D:\McoEasyTool\BatchFiles\RemoteBatchRunner.ps1 " + Server + " " + this.Target;
                    process.Start();
                    process.WaitForExit();
                    return true;
                }
                catch
                {
                    return false;
                }
            }

            public bool BatchStop(string Server)
            {
                return true;
            }

            public bool BatchRestart(string Server)
            {
                if (this.BatchStop(Server))
                {
                    return this.BatchStart(Server);
                }
                return false;
            }

            public string BatchGetState(string Server)
            {
                IntPtr userToken = IntPtr.Zero;
                bool success = McoUtilities.LogonUser(
                  HomeController.DEFAULT_USERNAME_IMPERSONNATION,
                  HomeController.DEFAULT_DOMAIN_IMPERSONNATION,
                  McoUtilities.Decrypt(HomeController.DEFAULT_PASSWORD_IMPERSONNATION),
                  (int)McoUtilities.LogonType.LOGON32_LOGON_INTERACTIVE, //2
                  (int)McoUtilities.LogonProvider.LOGON32_PROVIDER_DEFAULT, //0
                  out userToken);
                if (!success)
                {
                    return "Authentication Error";
                }
                using (WindowsIdentity.Impersonate(userToken))
                {
                    try
                    {
                        string ServerFolder = @"\\" + Server;
                        string FileName = "";
                        string[] ServerInfos = this.Target.Split(':');
                        if (ServerInfos[0].Length == 1 && ServerInfos.Length == 2)
                        {
                            string[] Path = ServerInfos[1].Split('\\');
                            ServerFolder += "\\" + ServerInfos[0] + "$";
                            for (int index = 0; index < Path.Length - 1; index++)
                            {
                                ServerFolder += "\\" + Path[index];
                            }
                            FileName = Path[Path.Length - 1];
                        }
                        string[] files = Directory.GetFiles(ServerFolder);
                        string alpha = ServerFolder + @"\" + FileName;
                        if (files.Contains(ServerFolder + @"\" + FileName))
                        {
                            return "Found";
                        }
                        else
                        {
                            return "Not found";
                        }
                    }
                    catch (Exception exception)
                    {
                        return "Exception " + exception.Message;
                    }
                }
            }
            //_______________________________________________________________________________

            //PROCESS FUNCTIONS
            public bool ProcessStart(string Server)
            {
                IntPtr userToken = IntPtr.Zero;
                bool success = McoUtilities.LogonUser(
                  HomeController.DEFAULT_USERNAME_IMPERSONNATION,
                  HomeController.DEFAULT_DOMAIN_IMPERSONNATION,
                  McoUtilities.Decrypt(HomeController.DEFAULT_PASSWORD_IMPERSONNATION),
                  (int)McoUtilities.LogonType.LOGON32_LOGON_INTERACTIVE, //2
                  (int)McoUtilities.LogonProvider.LOGON32_PROVIDER_DEFAULT, //0
                  out userToken);
                if (!success)
                {
                    return false;
                }
                using (WindowsIdentity.Impersonate(userToken))
                {
                    try
                    {
                        object[] theProcessToRun = { this.Target };
                        ConnectionOptions theConnection = new ConnectionOptions();
                        theConnection.Username = HomeController.DEFAULT_DOMAIN_IMPERSONNATION + "\\" + HomeController.DEFAULT_USERNAME_IMPERSONNATION;
                        theConnection.Password = McoUtilities.Decrypt(HomeController.DEFAULT_PASSWORD_IMPERSONNATION);
                        ManagementScope theScope = new ManagementScope("\\\\" + Server + "\\root\\cimv2", theConnection);
                        ManagementClass theClass = new ManagementClass(theScope, new ManagementPath("Win32_Process"), new ObjectGetOptions());
                        theClass.InvokeMethod("Create", theProcessToRun);
                        return true;
                    }
                    catch
                    {
                        return false;
                    }
                }
            }

            public bool ProcessStop(string Server)
            {
                IntPtr userToken = IntPtr.Zero;
                bool success = McoUtilities.LogonUser(
                  HomeController.DEFAULT_USERNAME_IMPERSONNATION,
                  HomeController.DEFAULT_DOMAIN_IMPERSONNATION,
                  McoUtilities.Decrypt(HomeController.DEFAULT_PASSWORD_IMPERSONNATION),
                  (int)McoUtilities.LogonType.LOGON32_LOGON_INTERACTIVE, //2
                  (int)McoUtilities.LogonProvider.LOGON32_PROVIDER_DEFAULT, //0
                  out userToken);
                if (!success)
                {
                    return false;
                }
                using (WindowsIdentity.Impersonate(userToken))
                {
                    Process[] processes = System.Diagnostics.Process.GetProcessesByName(this.Target, Server);
                    foreach (var process in processes)
                    {
                        try
                        {
                            process.Kill();
                        }
                        catch
                        {
                            return false;
                        }
                        return true;
                    }
                }
                return false;
            }

            public bool ProcessRestart(string Server)
            {
                if (this.ProcessStop(Server))
                {
                    return this.ProcessStart(Server);
                }
                return false;
            }

            public string ProcessGetState(string Server)
            {
                IntPtr userToken = IntPtr.Zero;
                bool success = McoUtilities.LogonUser(
                  HomeController.DEFAULT_USERNAME_IMPERSONNATION,
                  HomeController.DEFAULT_DOMAIN_IMPERSONNATION,
                  McoUtilities.Decrypt(HomeController.DEFAULT_PASSWORD_IMPERSONNATION),
                  (int)McoUtilities.LogonType.LOGON32_LOGON_INTERACTIVE, //2
                  (int)McoUtilities.LogonProvider.LOGON32_PROVIDER_DEFAULT, //0
                  out userToken);
                if (!success)
                {
                    return "Authentication Error";
                }
                using (WindowsIdentity.Impersonate(userToken))
                {
                    Process[] process = System.Diagnostics.Process.GetProcessesByName(this.Target.Split('.')[0], Server);
                    if (process.Length >= 1)
                    {
                        this.State = "Running";
                        return "Running";
                    }
                    else
                    {
                        this.State = "Stopped";
                        return "Stopped";
                    }
                }
            }
            //_______________________________________________________________________________

            public bool Start(string Server)
            {
                switch (this.Type)
                {
                    case "SERVICE": return this.ServiceStart(Server);
                    case "PROCESS": return this.ProcessStart(Server);
                    case "BATCH": return this.BatchStart(Server);
                }
                return false;
            }

            public bool Stop(string Server)
            {
                switch (this.Type)
                {
                    case "SERVICE": return this.ServiceStop(Server);
                    case "PROCESS": return this.ProcessStop(Server);
                    case "BATCH": return this.BatchStop(Server);
                }
                return false;
            }

            public bool Restart(string Server)
            {
                switch (this.Type)
                {
                    case "SERVICE": return this.ServiceRestart(Server);
                    case "PROCESS": return this.ProcessRestart(Server);
                    case "BATCH": return this.BatchRestart(Server);
                }
                return false;
            }

            public string GetState(string Server)
            {
                switch (this.Type)
                {
                    case "SERVICE": return this.ServiceGetState(Server);
                    case "PROCESS": return this.ProcessGetState(Server);
                    case "BATCH": return this.BatchGetState(Server);
                }
                return "Unknown";
            }

            public string Actionate(string Server)
            {
                switch (this.Action)
                {
                    case "START": return (this.Start(Server)) ? "Started" : "Failed";
                    case "STOP": return (this.Stop(Server)) ? "Stopped" : "Failed";
                    case "RESTART": return (this.Restart(Server)) ? "Restarted" : "Failed";
                    case "CHECK": return this.GetState(Server);
                    default: return "invalid action";
                }
            }

            public BrowseUrlResult CheckUrl(Selenium_Driver Navigator, AppServerInfo server)
            {
                BrowseUrlResult browse = new BrowseUrlResult(server);
                browse.Target += ":" + this.Target;
                if (this.Type != "URL")
                {
                    browse.Status = "KO";
                    browse.History = "";
                    browse.Details = "Not URL Type";
                    return browse;
                }
                browse.Status = "KO";
                browse.History += "Connexion\n";
                //CONNEXION TO URL
                try
                {
                    Navigator.driver.Navigate().GoToUrl(this.Target);
                    Navigator.WaitForPageLoad(25);
                    browse.Connexion = "OK";
                    browse.History += "Connecté\n";
                }
                catch (UnhandledAlertException)
                {
                    Navigator.driver.SwitchTo().Alert().Dismiss();
                }
                catch (OpenQA.Selenium.WebDriverTimeoutException)
                {
                    browse.Connexion = "Time out";
                    browse.Details = "URL: " + browse.Connexion;
                    return browse;
                }
                catch (Exception exception)
                {
                    browse.Connexion = exception.Message;
                    browse.Details = "URL: " + browse.Connexion;
                    return browse;
                }
                //CONNECTED TO URL

                //FILLING LOGIN FORM && AUTHENTICATION
                if (server.InputsParams.Count != 0)
                {
                    foreach (AppHtmlElementInfo input in server.InputsParams)
                    {
                        try
                        {
                            if (Navigator.SetNodeValue(input))
                            {
                                browse.ElementsResults[input] = "rempli";
                            }
                            else
                            {
                                browse.ElementsResults[input] = "non trouvé";
                            }
                            browse.History += input.ToString() + " " + browse.ElementsResults[input] + "\n";
                        }
                        catch (Exception exception)
                        {
                            browse.ElementsResults[input] = exception.Message;
                            browse.Details = "Auth: " + browse.ElementsResults[input];
                            return browse;
                        }
                    }

                    //AUTHENTICATION
                    try
                    {
                        IWebElement form = Navigator.GetNode(new AppHtmlElementInfo("FORM", "", "", "", "", ""));
                        form.Submit();
                        browse.History += "Authentification\n";
                    }
                    catch (Exception exception)
                    {
                        browse.Authentication = exception.Message;
                        browse.History += "Echec\n";
                        browse.Details = "Auth: " + browse.Authentication;
                        return browse;
                    }

                    try
                    {
                        Navigator.WaitForPageLoad(25);
                        browse.Authentication = "OK";
                        browse.History += "Authentifié\n";
                    }
                    catch (UnhandledAlertException)
                    {
                        Navigator.driver.SwitchTo().Alert().Accept();
                    }
                    catch (OpenQA.Selenium.WebDriverTimeoutException)
                    {
                        browse.Authentication = "Time out";
                        browse.Details = "Auth: " + browse.Authentication;
                        return browse;
                    }
                    catch (Exception exception)
                    {
                        browse.Authentication = exception.Message;
                        browse.Details = "Auth: " + browse.Authentication;
                        return browse;
                    }
                    //END AUTHENTICATION
                }

                //AUTHENTIFIED && CHECK VALUES
                if (server.OutputsParams.Count != 0)
                {
                    List<string> values = new List<string>();
                    foreach (AppHtmlElementInfo output in server.OutputsParams)
                    {
                        try
                        {
                            if (Navigator.isEqualNodeValue(output))
                            {
                                values.Add("OK");
                                browse.ElementsResults[output] = "OK";
                            }
                            else
                            {
                                values.Add("KO : " + output.ToString() + " : " + Navigator.GetNodeValue(output));
                                browse.ElementsResults[output] = "KO: " + Navigator.GetNodeValue(output);
                                browse.Values = output.ToString() + " " + browse.ElementsResults[output];
                            }
                            browse.History += browse.ElementsResults[output] + "\n";
                        }
                        catch (Exception exception)
                        {
                            values.Add("KO : " + output.ToString() + " : " + exception.Message);
                            browse.ElementsResults[output] = "KO: " + exception.Message;
                            browse.Values = output.ToString() + " " + browse.ElementsResults[output];
                            browse.Details = "Test: " + browse.Values;
                            return browse;
                        }
                    }
                    if (!values.Contains("KO"))
                    {
                        browse.Values = "OK";
                    }
                }
                bool con = false, auth = false, test = false;
                con = (browse.Connexion == "OK") ? true : false;
                auth = (browse.Authentication == "OK" || browse.Authentication == "NULL") ? true : false;
                test = (browse.Values == "OK" || browse.Values == "NULL") ? true : false;

                browse.Details = "Url:'" + browse.Connexion + "' Auth:'"
                    + browse.Authentication + "' Test:'" + browse.Values + "'";
                browse.Status = (con && auth && test) ? "OK" : "KO";
                browse.History += "Fin de traitement\n";
                return browse;
            }
        }

        public class AppServerInfo
        {

            public string Name { get; set; }
            public string Navigator { get; set; }
            public List<AppHtmlElementInfo> InputsParams { get; set; }
            public List<AppHtmlElementInfo> OutputsParams { get; set; }
            public List<Procedure> StartOrder { get; set; }
            public List<Procedure> StopOrder { get; set; }

            public AppServerInfo(AppServer server, Application application)
            {
                this.Name = server.Name;
                this.Navigator = application.Navigator;
                List<Procedure> Procedures = new List<Procedure>();
                List<Procedure> StopProcedures = new List<Procedure>();
                this.InputsParams = new List<AppHtmlElementInfo>();
                this.OutputsParams = new List<AppHtmlElementInfo>();
                try
                {
                    string[] StartOrder = server.StartOrder.Split(';');
                    for (int index = 0; index < StartOrder.Length; index++)
                    {
                        string[] infos = StartOrder[index].Split('|');
                        string type = infos[0];
                        string target = infos[1];
                        string action = infos[2];
                        Procedure procedure = new Procedure(target, type, action);
                        Procedures.Add(procedure);
                    }
                    this.StartOrder = Procedures;
                    if (server.StopOrder != "")
                    {
                        string[] StopOrder = server.StopOrder.Split(';');
                        for (int index = 0; index < StopOrder.Length; index++)
                        {
                            string[] infos = StopOrder[index].Split('|');
                            string type = infos[0];
                            string target = infos[1];
                            string action = infos[2];
                            Procedure procedure = new Procedure(target, type, action);
                            StopProcedures.Add(procedure);
                        }
                        this.StopOrder = StopProcedures;
                    }

                    foreach (AppHtmlElement element in application.AppHtmlElements)
                    {
                        if (element.Type == "INPUT")
                        {
                            this.InputsParams.Add(new AppHtmlElementInfo(element));
                        }
                        else
                        {
                            this.OutputsParams.Add(new AppHtmlElementInfo(element));
                        }
                    }
                }
                catch { }
            }

            public Dictionary<string, string> Start()
            {
                Dictionary<string, string> response = new Dictionary<string, string>();
                response.Add("status", "");
                response.Add("details", "");

                foreach (Procedure procedure in this.StartOrder)
                {
                    if (procedure.Type != "URL")
                    {
                        string result = procedure.Actionate(this.Name);
                        string comparator = "";
                        switch (procedure.Action)
                        {
                            case "START": comparator = "Started"; break;
                            case "STOP": comparator = "Stopped"; break;
                            case "RESTART": comparator = "Restarted"; break;
                            case "CHECK": comparator = (procedure.Type == "BATCH") ? "Found" : "Running"; break;
                            default: break;
                        }
                        response["details"] += procedure.Type + " " + procedure.Target + " " + procedure.Action + " : " + result + "\n";
                        if (result != comparator)
                        {
                            response["status"] = "KO";
                            return response;
                        }
                    }
                }
                response["status"] = "OK";
                return response;
            }

            public Dictionary<string, string> Stop()
            {
                Dictionary<string, string> response = new Dictionary<string, string>();
                response.Add("status", "");
                response.Add("details", "");

                foreach (Procedure procedure in this.StopOrder)
                {
                    if (procedure.Type != "URL")
                    {
                        string result = procedure.Actionate(this.Name);
                        string comparator = "";
                        switch (procedure.Action)
                        {
                            case "START": comparator = "Started"; break;
                            case "STOP": comparator = "Stopped"; break;
                            case "RESTART": comparator = "Restarted"; break;
                            case "CHECK": comparator = (procedure.Type == "BATCH") ? "Found" : "Running"; break;
                            default: break;
                        }
                        response["details"] += procedure.Type + " " + procedure.Target + " " + procedure.Action + " : " + result + "\n";
                        if (result != comparator)
                        {
                            response["status"] = "KO";
                            return response;
                        }
                    }
                }
                response["status"] = "OK";
                return response;
            }

            public Dictionary<string, string> Restart()
            {
                Dictionary<string, string> response = this.Stop();
                if (response["status"] == "OK")
                {
                    return this.Start();
                }
                else
                {
                    response["status"] = "KO";
                    return response;
                }
            }

            public Dictionary<string, string> GetState()
            {
                Dictionary<string, string> response = new Dictionary<string, string>();
                response.Add("status", "");
                response.Add("details", "");

                foreach (Procedure procedure in this.StartOrder)
                {
                    if (procedure.Type != "URL")
                    {
                        string result = procedure.GetState(this.Name);
                        string comparator = "";//(procedure.Type == "BATCH") ? "Found" : "Running";
                        if (procedure.Action == "START" || procedure.Action == "RESTART" || procedure.Action == "CHECK")
                        {
                            comparator = (procedure.Type == "BATCH") ? "Found" : "Running";
                        }
                        else
                        {
                            comparator = (procedure.Type == "BATCH") ? "Found" : "Stopped";
                        }
                        response["details"] += procedure.Type + " " + procedure.Target + " : " + result + "\n";
                        if (result != comparator)
                        {
                            response["status"] = "KO";
                            return response;
                        }
                    }
                }
                response["status"] = "OK";
                return response;
            }

            public static Dictionary<AppServerInfo, List<BrowseUrlResult>> BrowseServerInfos(List<AppServerInfo> servers)
            {
                Dictionary<AppServerInfo, List<BrowseUrlResult>> Results = new Dictionary<AppServerInfo, List<BrowseUrlResult>>();
                if (servers.Count != 0)
                {
                    Selenium_Driver Navigator = new Selenium_Driver();
                    foreach (AppServerInfo server in servers)
                    {
                        try
                        {
                            List<Procedure> procedures = server.StartOrder.Where(pro => pro.Type == "URL").ToList();
                            if (procedures.Count != 0)
                            {
                                List<BrowseUrlResult> browses = new List<BrowseUrlResult>();
                                foreach (Procedure procedure in procedures)
                                {
                                    Navigator.Setup(server.Navigator);
                                    BrowseUrlResult browse = procedure.CheckUrl(Navigator, server);
                                    browses.Add(browse);
                                    Navigator.Teardown(server.Navigator);
                                }

                                Results.Add(server, browses);
                            }
                        }
                        catch
                        {
                            List<BrowseUrlResult> browses = new List<BrowseUrlResult>();
                            Results.Add(server, browses);
                            Navigator.Teardown(server.Navigator);
                            continue;
                        }
                    }
                }
                return Results;
            }

            public AppServer GetServer(Dictionary<AppServer, List<BrowseUrlResult>> UrlBrowseResults)
            {
                foreach (KeyValuePair<AppServer, List<BrowseUrlResult>> result in UrlBrowseResults)
                {
                    if (result.Key.Name == this.Name && result.Key.Application.Navigator == this.Navigator
                        && result.Key.Application.AppHtmlElements.Count == (this.InputsParams.Count + this.OutputsParams.Count)
                      )
                    {
                        return result.Key;
                    }
                }
                return null;
            }
        }

        public class ApplicationInfo
        {
            public string Name { get; set; }
            public string Url { get; set; }
            public string Navigator { get; set; }
            public List<AppHtmlElementInfo> InputsParams { get; set; }
            public List<AppHtmlElementInfo> OutputsParams { get; set; }
            public AppHtmlElementInfo LoginButton { get; set; }
            public List<AppServerInfo> Servers { get; set; }

            public ApplicationInfo(Application application)
            {
                try
                {
                    this.Name = application.Name;
                    this.Url = application.Url;
                    this.Navigator = application.Navigator;
                    List<AppHtmlElement> elements = application.AppHtmlElements.ToList();
                    this.InputsParams = new List<AppHtmlElementInfo>();
                    this.OutputsParams = new List<AppHtmlElementInfo>();
                    this.LoginButton = null;
                    foreach (AppHtmlElement element in elements)
                    {
                        switch (element.Type)
                        {
                            case "INPUT":
                                this.InputsParams.Add(new AppHtmlElementInfo(element));
                                break;
                            case "OUTPUT":
                                this.OutputsParams.Add(new AppHtmlElementInfo(element));
                                break;
                            case "LOGIN":
                                if (this.LoginButton == null)
                                {
                                    LoginButton = new AppHtmlElementInfo(element);
                                }
                                break;
                            default: break;
                        }
                    }
                    List<AppServerInfo> ServersInfo = new List<AppServerInfo>();
                    AppServer[] servers = application.AppServers.ToArray();
                    foreach (AppServer server in servers)
                    {
                        AppServerInfo serverInfo = new AppServerInfo(server, server.Application);
                        ServersInfo.Add(serverInfo);
                    }
                    this.Servers = ServersInfo;
                }
                catch { }
            }

            public Dictionary<string, string> Start()
            {
                Dictionary<string, string> response = new Dictionary<string, string>();
                response.Add("status", "OK");
                response.Add("details", "");
                List<string> Status = new List<string>();
                foreach (AppServerInfo server in this.Servers)
                {
                    Dictionary<string, string> result = server.Start();
                    response["details"] += "*****************************************\n" +
                        server.Name + " : " + result["status"] + "\n" + result["details"];
                    response["details"] += "*****************************************\n";
                    if (result["status"] != "OK")
                    {
                        Status.Add("KO");
                    }
                    else
                    {
                        Status.Add("OK");
                    }
                }
                if (!Status.Contains("KO"))
                {
                    response["status"] = "OK";
                }
                else
                {
                    if (!Status.Contains("OK"))
                    {
                        response["status"] = "KO";
                    }
                    else
                    {
                        response["status"] = "H-OK";
                    }
                }
                return response;
            }

            public Dictionary<string, string> Stop()
            {
                Dictionary<string, string> response = new Dictionary<string, string>();
                response.Add("status", "OK");
                response.Add("details", "");
                List<string> Status = new List<string>();
                foreach (AppServerInfo server in this.Servers)
                {
                    Dictionary<string, string> result = server.Stop();
                    response["details"] += "*****************************************\n" +
                        server.Name + " : " + result["status"] + "\n" + result["details"];
                    response["details"] += "*****************************************\n";
                    if (result["status"] != "OK")
                    {
                        Status.Add("KO");
                    }
                    else
                    {
                        Status.Add("OK");
                    }
                }
                if (!Status.Contains("KO"))
                {
                    response["status"] = "OK";
                }
                else
                {
                    if (!Status.Contains("OK"))
                    {
                        response["status"] = "KO";
                    }
                    else
                    {
                        response["status"] = "H-OK";
                    }
                }
                return response;
            }

            public Dictionary<string, string> Restart()
            {
                Dictionary<string, string> response = this.Stop();
                if (response["status"] == "OK")
                {
                    return this.Start();
                }
                else
                {
                    response["status"] = "KO";
                    return response;
                }
            }

            public Dictionary<string, string> GetState()
            {
                Dictionary<string, string> response = new Dictionary<string, string>();
                response.Add("status", "");
                response.Add("details", "");
                List<string> Status = new List<string>();
                foreach (AppServerInfo server in this.Servers)
                {
                    Dictionary<string, string> result = server.GetState();
                    response["details"] += "*****************************************\nServeur " +
                        server.Name + " : " + result["status"] + "\n" + result["details"];
                    response["details"] += "*****************************************\n";
                    if (result["status"] != "OK")
                    {
                        Status.Add("KO");
                    }
                    else
                    {
                        Status.Add("OK");
                    }
                }
                if (!Status.Contains("KO"))
                {
                    response["status"] = "OK";
                }
                else
                {
                    if (!Status.Contains("OK"))
                    {
                        response["status"] = "KO";
                    }
                    else
                    {
                        response["status"] = "H-OK";
                    }
                }
                return response;
            }

            public string TestUrl()
            {
                try
                {
                    if (this.Url != null && this.Url.Trim() != "")
                    {
                        string username = HomeController.DEFAULT_USERNAME_IMPERSONNATION,
                            domain = HomeController.DEFAULT_DOMAIN_IMPERSONNATION, password = McoUtilities.Decrypt(HomeController.DEFAULT_PASSWORD_IMPERSONNATION);
                        CredentialCache credentialCache = new CredentialCache();
                        NetworkCredential credentials = new NetworkCredential(username, password, domain);
                        credentialCache.Add(new System.Uri(this.Url), "Ntlm", credentials);
                        HttpWebRequest httpWebRequest = (HttpWebRequest)WebRequest.Create((this.Url));
                        httpWebRequest.Credentials = credentialCache;
                        httpWebRequest.PreAuthenticate = true;
                        HttpWebRequest request = (HttpWebRequest)WebRequest.Create(this.Url);
                        request.Credentials = credentialCache;
                        HttpWebResponse answer = (HttpWebResponse)request.GetResponse();
                        if (answer.StatusCode.ToString() == "OK")
                        {
                            return "La page a répondu avec le code : " + answer.StatusCode.ToString();
                        }
                        else
                        {
                            IntPtr userToken = IntPtr.Zero;
                            bool success = McoUtilities.LogonUser(
                                username,
                                domain,
                                password,
                                (int)McoUtilities.LogonType.LOGON32_LOGON_INTERACTIVE, //2
                            (int)McoUtilities.LogonProvider.LOGON32_PROVIDER_DEFAULT, //0
                            out userToken);
                            using (WindowsIdentity.Impersonate(userToken))
                            {
                                HttpWebResponse anaswer = (HttpWebResponse)request.GetResponse();
                                if (anaswer.StatusCode.ToString() == "OK")
                                {
                                    return "La page a répondu avec le code : " + anaswer.StatusCode.ToString();
                                }
                                else
                                {
                                    return "La page a répondu avec le code : " + anaswer.StatusCode.ToString() + "\n" + anaswer.StatusDescription;
                                }
                            }
                        }
                    }
                    else
                    {
                        return "Pas d'Url fournie";
                    }
                }
                catch (Exception exception)
                {
                    return "Une exception est survenue : " + exception.Message;
                }
            }

            public BrowseUrlResult FillInputs(Selenium_Driver Navigator, BrowseUrlResult browse)
            {
                if (this.InputsParams.Count != 0)
                {
                    foreach (AppHtmlElementInfo input in this.InputsParams)
                    {
                        try
                        {
                            if (Navigator.SetNodeValue(input))
                            {
                                browse.ElementsResults[input] = "rempli";
                            }
                            else
                            {
                                browse.ElementsResults[input] = "non trouvé";
                            }
                            browse.History += input.ToString() + " " + browse.ElementsResults[input] + "\n";
                        }
                        catch (Exception exception)
                        {
                            browse.ElementsResults[input] = exception.Message;
                            browse.Details = "Auth: " + browse.ElementsResults[input];
                        }
                    }
                }
                return browse;
            }

            public BrowseUrlResult CheckOutputs(Selenium_Driver Navigator, BrowseUrlResult browse)
            {
                if (this.OutputsParams.Count != 0)
                {
                    List<string> values = new List<string>();
                    foreach (AppHtmlElementInfo output in this.OutputsParams)
                    {
                        try
                        {
                            if (Navigator.isEqualNodeValue(output))
                            {
                                values.Add("OK");
                                browse.ElementsResults[output] = "OK";
                            }
                            else
                            {
                                values.Add("KO : " + output.ToString() + " : " + Navigator.GetNodeValue(output));
                                browse.ElementsResults[output] = "KO: " + Navigator.GetNodeValue(output);
                                browse.Values = output.ToString() + " " + browse.ElementsResults[output];
                            }
                            browse.History += browse.ElementsResults[output] + "\n";
                        }
                        catch (Exception exception)
                        {
                            values.Add("KO : " + output.ToString() + " : " + exception.Message);
                            browse.ElementsResults[output] = "KO: " + exception.Message;
                            browse.Values = output.ToString() + " " + browse.ElementsResults[output];
                            browse.Details = "Test: " + browse.Values;
                            return browse;
                        }
                    }
                    bool faults = false;
                    foreach (string value in values)
                    {
                        if (value.IndexOf("KO: ") != -1)
                        {
                            faults = true;
                            break;
                        }
                    }
                    if (!faults)
                    {
                        browse.Values = "OK";
                    }
                }
                return browse;
            }

            public BrowseUrlResult Authenticate(Selenium_Driver Navigator, BrowseUrlResult browse, bool force = false)
            {
                //AUTHENTICATION
                bool posted = false;
                int timeout = 35;
                try
                {
                    if (this.LoginButton != null)
                    {
                        IWebElement button = Navigator.GetNode(this.LoginButton);
                        if (button != null)
                        {
                            timeout = 40;
                            button.Click();
                            posted = true;
                        }
                        else
                        {
                            button = Navigator.GetNodeByContent(this.LoginButton);
                            if (button != null)
                            {
                                timeout = 40;
                                button.Click();
                                posted = true;
                            }
                        }
                    }
                    else
                    {
                        if (force)
                        {
                            timeout = 40;
                            posted = Navigator.SubmitForm(1);
                        }
                        else
                        {
                            posted = Navigator.SubmitForm(0);
                        }
                    }
                    browse.History += (posted) ? "Authentification\n" : "Pas d'Authentification\n";
                }
                catch (Exception exception)
                {
                    browse.Authentication = exception.Message;
                    browse.History += "Echec\n";
                    browse.Details = "Auth: " + browse.Authentication;
                    return browse;
                }

                try
                {
                    Navigator.WaitForPageLoad(timeout);
                    browse.Authentication = "OK";
                    browse.History += "Authentifié\n";
                }
                catch (UnhandledAlertException)
                {
                    Navigator.driver.SwitchTo().Alert().Accept();
                }
                catch (OpenQA.Selenium.WebDriverTimeoutException)
                {
                    browse.Authentication = "Time out";
                    browse.Details = "Auth: " + browse.Authentication;
                    return browse;
                }
                catch (Exception exception)
                {
                    browse.Authentication = exception.Message;
                    browse.Details = "Auth: " + browse.Authentication;
                    return browse;
                }
                //END AUTHENTICATION
                return browse;
            }

            public bool isAuthenticated(Selenium_Driver Navigator, int previous_NodesNumber, string previous_url,
                string previous_title, string previous_source, List<string> previous_handles = null)
            {
                //FIRST TEST Login nodes presence
                List<bool> present_inputs = new List<bool>();
                foreach (AppHtmlElementInfo input in this.InputsParams)
                {
                    try
                    {
                        if (Navigator.GetNode(input) != null)
                        {
                            present_inputs.Add(true);
                        }
                        else
                        {
                            present_inputs.Add(false);
                        }
                    }
                    catch
                    {
                        present_inputs.Add(false);
                    }
                }
                if (this.LoginButton != null)
                {
                    if (Navigator.GetNode(this.LoginButton) != null)
                    {
                        present_inputs.Add(true);
                    }
                    else
                    {
                        present_inputs.Add(false);
                    }
                }
                if (!present_inputs.Contains(false))
                {
                    //ALL THE NODES ARE REPEATED ON NEW PAGE, OTHER CHECKS ARE REQUIRED
                    int NodesNumber = Navigator.driver.FindElements(By.TagName("*")).Count;
                    if (NodesNumber == previous_NodesNumber &&
                        Navigator.driver.Url == previous_url &&
                        previous_title == Navigator.driver.Title &&
                        previous_source == Navigator.driver.PageSource)
                    { return false; }
                }
                return true;
            }

            public bool TestLeaving(Selenium_Driver Navigator)
            {
                if (this.InputsParams.Count != 0)
                {
                    return !(Navigator.GetNode(this.InputsParams.FirstOrDefault()).Displayed);
                }
                if (this.LoginButton != null)
                {
                    return !(Navigator.GetNode(this.LoginButton).Displayed);
                }
                if (this.OutputsParams.Count != 0)
                {
                    return !(Navigator.GetNode(OutputsParams.FirstOrDefault()).Displayed);
                }
                return true;
            }

            public BrowseUrlResult Browse(Selenium_Driver Navigator)
            {
                BrowseUrlResult browse = new BrowseUrlResult(this);
                browse.Status = "KO";
                browse.History += "Connexion\n";

                //CONNEXION TO URL
                try
                {
                    Navigator.driver.Navigate().GoToUrl(this.Url);
                    Navigator.WaitForPageLoad(15);
                    browse.Connexion = "OK";
                    browse.History += "Connecté\n";
                }
                catch (UnhandledAlertException)
                {
                    Navigator.driver.SwitchTo().Alert().Dismiss();
                }
                catch (OpenQA.Selenium.WebDriverTimeoutException)
                {
                    if (TestLeaving(Navigator))
                    {
                        browse.Connexion = "Time out";
                        browse.Details = "URL: " + browse.Connexion;
                        return browse;
                    }
                }
                catch (Exception exception)
                {
                    if (TestLeaving(Navigator))
                    {
                        browse.Connexion = exception.Message;
                        browse.Details = "URL: " + browse.Connexion;
                        return browse;
                    }
                }
                //CONNECTED TO URL

                //FILLING LOGIN FORM && AUTHENTICATION
                if (this.InputsParams.Count != 0)
                {
                    browse = FillInputs(Navigator, browse);
                    int previous_NodesNumber = Navigator.driver.FindElements(By.TagName("*")).Count;
                    string previous_url = Navigator.driver.Url, previous_title = Navigator.driver.Title;
                    string previous_source = Navigator.driver.PageSource;
                    List<string> previous_handles = Navigator.driver.WindowHandles.ToList();

                    browse = Authenticate(Navigator, browse);
                    //CHECK THAT WE DON'T POSSESS OBJECTS FROM LOGIN PAGE
                    //CHECK AUTHENTICATION
                    if (!isAuthenticated(Navigator, previous_NodesNumber, previous_url, previous_title, previous_source))
                    {
                        FillInputs(Navigator, browse);
                        Authenticate(Navigator, browse, true);

                        //CHECK AGAIN
                        if (!isAuthenticated(Navigator, previous_NodesNumber, previous_url, previous_title, previous_source))
                        {
                            browse.Authentication = "Erreur Authentification";
                            browse.Details = "Auth: " + browse.Authentication;
                            return browse;
                        }
                    }
                    //END CHECK
                }

                //AUTHENTIFIED && CHECK VALUES
                if (this.OutputsParams.Count != 0)
                {
                    browse = CheckOutputs(Navigator, browse);
                }
                //END CHECK

                //CONCLUDE APPLICATION STATE
                bool con = false, auth = false, test = false;
                con = (browse.Connexion == "OK") ? true : false;
                auth = (browse.Authentication == "OK" || browse.Authentication == "NULL") ? true : false;
                test = (browse.Values == "OK" || browse.Values == "NULL") ? true : false;

                browse.Details = "Url:'" + browse.Connexion + "' Auth:'"
                    + browse.Authentication + "' Test:'" + browse.Values + "'";
                browse.Status = (con && auth && test) ? "OK" : "KO";
                browse.History += "Fin de traitement\n";
                return browse;
            }

            public Dictionary<ApplicationInfo, BrowseUrlResult> BrowseApplicationInfo()
            {
                Dictionary<ApplicationInfo, BrowseUrlResult> Results = new Dictionary<ApplicationInfo, BrowseUrlResult>();
                Selenium_Driver Navigator = new Selenium_Driver();
                try
                {
                    Navigator.Setup(this.Navigator);
                    BrowseUrlResult browse = this.Browse(Navigator);
                    Results.Add(this, browse);
                    Navigator.Teardown(this.Navigator);
                }
                catch
                {
                    BrowseUrlResult browse = new BrowseUrlResult();
                    Results.Add(this, browse);
                    Navigator.Teardown(this.Navigator);
                }
                return Results;
            }

            public static Dictionary<ApplicationInfo, BrowseUrlResult> BrowseApplicationInfos(List<ApplicationInfo> applications)
            {
                Dictionary<ApplicationInfo, BrowseUrlResult> Results = new Dictionary<ApplicationInfo, BrowseUrlResult>();
                if (applications.Count != 0)
                {
                    foreach (ApplicationInfo application in applications)
                    {
                        try
                        {
                            Dictionary<ApplicationInfo, BrowseUrlResult> Result = new Dictionary<ApplicationInfo, BrowseUrlResult>();
                            Result = application.BrowseApplicationInfo();
                            if (Result != null && Result.Count != 0)
                            {
                                Results.Add(application, Result[application]);
                            }
                            else
                            {
                                BrowseUrlResult browse = new BrowseUrlResult();
                                Result.Add(application, browse);
                            }
                        }
                        catch
                        {
                            BrowseUrlResult browse = new BrowseUrlResult();
                            Results.Add(application, browse);
                            continue;
                        }
                    }
                }
                return Results;
            }

            public Application GetApplication(Dictionary<Application, BrowseUrlResult> ApplicationBrowseResults)
            {
                foreach (KeyValuePair<Application, BrowseUrlResult> result in ApplicationBrowseResults)
                {
                    int number = (this.LoginButton != null) ? 1 : 0;
                    if (result.Key.Name == this.Name && result.Key.Url == this.Url
                        && result.Key.AppServers.Count == this.Servers.Count
                        && result.Key.AppHtmlElements.Count == (this.OutputsParams.Count + this.InputsParams.Count + number)
                      )
                    {
                        return result.Key;
                    }
                }
                return null;
            }
        }

        public class AppHtmlElementInfo
        {
            public string TagName { get; set; }
            public string AttrName { get; set; }
            public string AttrId { get; set; }
            public string AttrClass { get; set; }
            public string AttrXpath { get; set; }
            public string Value { get; set; }
            public string Type { get; set; }

            public AppHtmlElementInfo(AppHtmlElement htmlelement)
            {
                this.TagName = htmlelement.TagName;
                this.AttrName = htmlelement.AttrName;
                this.AttrId = htmlelement.AttrId;
                this.AttrClass = htmlelement.AttrClass;
                this.AttrXpath = htmlelement.AttrXpath;
                this.Value = htmlelement.Value;
                this.Type = htmlelement.Type;
            }

            public AppHtmlElementInfo(string TagName, string Name, string Id, string Class, string Xpath, string Value)
            {
                this.TagName = TagName;
                this.AttrName = Name;
                this.AttrId = Id;
                this.AttrClass = Class;
                this.AttrXpath = Xpath;
                this.Value = Value;
                this.Type = "";
            }

            public string GetStrongerSelector()
            {
                if (this.AttrXpath != null && this.AttrXpath.Trim() != "")
                {
                    return "XPATH";
                }
                if (this.AttrId != null && this.AttrId.Trim() != "")
                {
                    return "ID";
                }
                if (this.AttrName != null && this.AttrName.Trim() != "")
                {
                    return "NAME";
                }
                if (this.AttrClass != null && this.AttrClass.Trim() != "")
                {
                    return "CLASS";
                }
                return "NONE";
            }

            public override string ToString()
            {
                return "Balise " + this.TagName + " id='" + this.AttrId + "' Name='" + this.AttrName + "' Class='" + this.AttrClass + "' Value='" + this.Value + "'";
            }
        }

        public class BrowseUrlResult
        {
            public Dictionary<AppHtmlElementInfo, string> ElementsResults { get; set; }
            public string Target { get; set; }
            public string Status { get; set; }
            public string Details { get; set; }
            public string History { get; set; }
            public string Connexion { get; set; }
            public string Authentication { get; set; }
            public string Values { get; set; }

            public BrowseUrlResult()
            {
                this.ElementsResults = new Dictionary<AppHtmlElementInfo, string>();
                this.Status = this.Details = this.History = this.Target = "";
                this.Connexion = this.Authentication = this.Values = "";
            }

            public BrowseUrlResult(Application application)
            {
                this.Status = this.Details = this.History = "";
                this.Connexion = "";
                this.Target = application.Name;
                this.ElementsResults = new Dictionary<AppHtmlElementInfo, string>();
                foreach (AppHtmlElement element in application.AppHtmlElements)
                {
                    ElementsResults.Add(new AppHtmlElementInfo(element), "");
                }
                if (application.AppHtmlElements.Where(input => input.Type == "INPUT").Count() == 0)
                {
                    this.Authentication = "NULL";
                }
                if (application.AppHtmlElements.Where(output => output.Type == "OUTPUT").Count() == 0)
                {
                    this.Values = "NULL";
                }
            }

            public BrowseUrlResult(ApplicationInfo application)
            {
                this.Status = this.Details = this.History = "";
                this.Connexion = "";
                this.Target = application.Name;
                this.ElementsResults = new Dictionary<AppHtmlElementInfo, string>();
                foreach (AppHtmlElementInfo element in application.InputsParams)
                {
                    ElementsResults.Add(element, "");
                }
                foreach (AppHtmlElementInfo element in application.OutputsParams)
                {
                    ElementsResults.Add(element, "");
                }

                if (application.InputsParams.Count() == 0)
                {
                    this.Authentication = "NULL";
                }
                if (application.OutputsParams.Count() == 0)
                {
                    this.Values = "NULL";
                }
            }

            public BrowseUrlResult(AppServer server)
            {
                this.Status = this.Details = this.History = "";
                this.Connexion = "";
                this.Target = server.Name;
                this.ElementsResults = new Dictionary<AppHtmlElementInfo, string>();
                foreach (AppHtmlElement element in server.Application.AppHtmlElements)
                {
                    ElementsResults.Add(new AppHtmlElementInfo(element), "");
                }
                if (server.Application.AppHtmlElements.Where(input => input.Type == "INPUT").Count() == 0)
                {
                    this.Authentication = "NULL";
                }
                if (server.Application.AppHtmlElements.Where(output => output.Type == "OUTPUT").Count() == 0)
                {
                    this.Values = "NULL";
                }
            }

            public BrowseUrlResult(AppServerInfo server)
            {
                this.Status = this.Details = this.History = "";
                this.Connexion = "";
                this.Target = server.Name;
                this.ElementsResults = new Dictionary<AppHtmlElementInfo, string>();
                foreach (AppHtmlElementInfo element in server.InputsParams)
                {
                    ElementsResults.Add(element, "");
                }
                foreach (AppHtmlElementInfo element in server.OutputsParams)
                {
                    ElementsResults.Add(element, "");
                }

                if (server.InputsParams.Count() == 0)
                {
                    this.Authentication = "NULL";
                }
                if (server.OutputsParams.Count() == 0)
                {
                    this.Values = "NULL";
                }
            }
        }

        [HttpPost]
        public JsonResult AddApplication()
        {
            string Name = "", Url = "", Domain = "", Navigator = "IE";
            Dictionary<string, string> response = new Dictionary<string, string>();
            response.Add("status", "");
            response.Add("appId", "");
            try
            {
                Name = Request.Form["Name"].ToString();
                Url = Request.Form["Url"].ToString();
                Domain = Request.Form["Domain"].ToString();
                Navigator = Request.Form["Navigator"].ToString();

                int Id = 0; Int32.TryParse(Domain, out Id);
                AppDomain domain = db.AppDomains.Find(Id);
                if (domain == null)
                {
                    response["status"] = "Erreur lors de la récupération du domaine d'application";
                    return Json(response, JsonRequestBehavior.AllowGet);
                }
                Application application = db.Applications.Create();
                application.Name = Name;
                application.Url = Url;
                application.Domain = domain.Id;
                application.Navigator = (Navigator.Trim() != "") ? Navigator : "IE";
                if (ModelState.IsValid)
                {
                    db.Applications.Add(application);
                    db.SaveChanges();
                    domain.Applications = db.Applications.Where(app => app.Domain == domain.Id).Count();
                    db.Entry(domain).State = System.Data.Entity.EntityState.Modified;
                    db.SaveChanges();
                    response["status"] = "OK";
                    response["appId"] = application.Id.ToString();
                    Specific_Logging(new Exception(""), "AddApplication " + application.Name, 3);
                    return Json(response, JsonRequestBehavior.AllowGet);
                }
                response["status"] = "KO";
                Specific_Logging(new Exception(""), "AddApplication " + application.Name, 2);
                return Json(response, JsonRequestBehavior.AllowGet);
            }
            catch (Exception exception)
            {
                Specific_Logging(exception, "AddApplication");
                response["status"] = "Erreur lors de l'ajout de l'application";
                return Json(response, JsonRequestBehavior.AllowGet);
            }
        }

        public string GetServersServices()
        {
            string[] Servers = null;
            string Services = "";
            IntPtr userToken = IntPtr.Zero;

            List<string> Services_Names = new List<string>();
            List<string> Display_Names = new List<string>();
            List<string> Services_list = new List<string>();
            try
            {
                Servers = Request.Form["Servers"].ToString().Split(';');
                int first = 0;
                foreach (string server in Servers)
                {
                    if (server.Trim() == "")
                    {
                        continue;
                    }
                    List<string> newServiceNamelist = new List<string>();
                    List<string> newDisplayNamelist = new List<string>();
                    bool success = McoUtilities.LogonUser(
                        HomeController.DEFAULT_USERNAME_IMPERSONNATION,
                        HomeController.DEFAULT_DOMAIN_IMPERSONNATION,
                        McoUtilities.Decrypt(HomeController.DEFAULT_PASSWORD_IMPERSONNATION),
                        (int)McoUtilities.LogonType.LOGON32_LOGON_INTERACTIVE, //2
                        (int)McoUtilities.LogonProvider.LOGON32_PROVIDER_DEFAULT, //0
                        out userToken);
                    using (WindowsIdentity.Impersonate(userToken))
                    {
                        ServiceController[] services = ServiceController.GetServices(server.Trim());
                        foreach (ServiceController service in services)
                        {
                            newServiceNamelist.Add(service.ServiceName.ToLower());
                            newDisplayNamelist.Add(service.DisplayName.ToLower());
                        }
                    }
                    if (first == 0)
                    {
                        Services_Names = newServiceNamelist;
                        Display_Names = newDisplayNamelist;
                        //Services_list = newlist;
                    }
                    else
                    {
                        Services_Names = Services_Names.Intersect(newServiceNamelist).ToList();
                        Display_Names = Display_Names.Intersect(newDisplayNamelist).ToList();
                        //Services_list = Services_list.Intersect(newlist).ToList();
                    }
                    first++;
                }
                if (Display_Names.Count >= Services_Names.Count)
                {
                    Services_list = Display_Names;
                }
                else
                {
                    Services_list = Services_Names;
                }
                foreach (string Service in Services_list.OrderBy(name => name).ToList())
                {
                    Services += Service + "; ";
                }
                Services = Services.Substring(0, Services.Length - 2);
                return Services;
            }
            catch
            {
                return Services;
            }
        }

        public string ApplicationExists()
        {
            try
            {
                string Name = Request.Form["Name"].ToString();
                List<Application> applications = db.Applications.ToList();
                foreach (Application application in applications)
                {
                    if (application.Name.ToLower() == Name.ToLower())
                    {
                        return "KO";
                    }
                }
                return "OK";
            }
            catch
            {
                return "KO";
            }
        }

        public string DeleteApplication(int id)
        {
            Application application = db.Applications.Find(id);

            AppServer[] servers = application.AppServers.ToArray();
            AppHtmlElement[] htmelelements = application.AppHtmlElements.ToArray();
            Application_Report[] appreports = application.Application_Reports.ToArray();
            Scheduled_Application[] scheduledapplications = application.Scheduled_Applications.ToArray();
            foreach (Application_Report appreport in appreports)
            {
                AppReport report = appreport.AppReport;
                DeleteAppReport(report.Id);
            }
            foreach (AppServer server in servers)
            {
                db.AppServers.Remove(server);
                db.SaveChanges();
            }
            foreach (AppHtmlElement htmlelement in htmelelements)
            {
                db.AppHtmlElements.Remove(htmlelement);
                db.SaveChanges();
            }
            foreach (Scheduled_Application scheduledapplication in scheduledapplications)
            {
                db.ScheduledApplications.Remove(scheduledapplication);
                db.SaveChanges();
            }

            db.Applications.Remove(application);
            db.SaveChanges();
            Specific_Logging(new Exception(""), "DeleteApplication " + application.Name, 3);
            return "L'application " + application.Name + " a été supprimée";
        }

        [HttpPost]
        public string EditApplication(int id)
        {
            try
            {
                string Name = Request.Form["Name"].ToString();
                string Url = Request.Form["Url"].ToString();
                string Domain = Request.Form["Domain"].ToString();
                string Navigator = Request.Form["Navigator"].ToString();
                int Id = 0; Int32.TryParse(Domain, out Id);
                AppDomain domain = db.AppDomains.Find(Id);
                if (domain == null)
                {
                    return "Erreur lors de la récupération du domaine d'application";
                }
                Application application = db.Applications.Find(id);
                application.Name = Name;
                application.Url = Url;
                application.Domain = domain.Id;
                application.Navigator = (Navigator.Trim() != "") ? Navigator : "IE";
                if (ModelState.IsValid)
                {
                    db.Entry(application).State = System.Data.Entity.EntityState.Modified;
                    db.SaveChanges();
                    domain.Applications = db.Applications.Where(app => app.Domain == domain.Id).Count();
                    db.Entry(domain).State = System.Data.Entity.EntityState.Modified;
                    db.SaveChanges();
                    Specific_Logging(new Exception(""), "EditApplication " + application.Name, 3);
                    return "OK";
                }
                Specific_Logging(new Exception(""), "EditApplication " + application.Name, 2);
                return "KO";
            }
            catch (Exception exception)
            {
                Specific_Logging(exception, "AddApplication");
                return "Erreur lors de la modification de l'application";
            }
        }

        [HttpPost]
        public string StopEditApplication(int id)
        {
            try
            {
                Application application = db.Applications.Find(id);
                string Name = Request.Form["Name"].ToString();
                string Lines = Request.Form["Procedures"].ToString();
                application.Name = Name;
                if (ModelState.IsValid)
                {
                    db.Entry(application).State = System.Data.Entity.EntityState.Modified;
                    db.SaveChanges();
                    Specific_Logging(new Exception(""), "StopEditApplication " + application.Name, 3);
                    return "OK";
                }
                return "KO";
            }
            catch (Exception exception)
            {
                Specific_Logging(exception, "SopEditApplication");
                return "Erreur lors de la modification de l'application";
            }
        }

        public string TestApplicationUrl(int id)
        {
            Application application = db.Applications.Find(id);
            if (application == null)
            {
                return "Inconnu";
            }
            try
            {
                ApplicationInfo app = new ApplicationInfo(application);
                string linkable = app.TestUrl();
                return linkable;
            }
            catch (Exception exception)
            {
                Specific_Logging(exception, "TestApplicationUrl " + application.Name, 3);
                return "Inconnu";
            }
        }

        public JsonResult GetApplicationState(int id)
        {
            Dictionary<string, string> response = new Dictionary<string, string>();
            response.Add("status", "");
            response.Add("details", "");
            Application application = db.Applications.Find(id);
            if (application == null)
            {
                response["status"] = "KO";
                response["details"] = "L'application n'a pas été retrouvée dans la base de données.";
                return Json(response, JsonRequestBehavior.AllowGet);
            }
            try
            {
                ApplicationInfo app = new ApplicationInfo(application);
                return Json(app.GetState(), JsonRequestBehavior.AllowGet);
            }
            catch (Exception exception)
            {
                Specific_Logging(exception, "GetApplicationState");
                response["status"] = "KO";
                response["details"] = "Exception " + exception.Message;
                return Json(response, JsonRequestBehavior.AllowGet);
            }
        }

        public JsonResult StartApplication(int id)
        {
            Dictionary<string, string> response = new Dictionary<string, string>();
            response.Add("status", "");
            response.Add("details", "");
            Application application = db.Applications.Find(id);
            if (application == null)
            {
                response["status"] = "KO";
                response["details"] = "L'application n'a pas été démarrée car elle n'a pas été retrouvée dans la base de données.";
                return Json(response, JsonRequestBehavior.AllowGet);
            }
            try
            {
                ApplicationInfo app = new ApplicationInfo(application);
                Dictionary<string, string> status = app.GetState();
                if (status["status"] == "OK")
                {
                    response["status"] = "OK";
                    response["details"] = "L'application est déjà en marche.";
                }
                else
                {
                    Specific_Logging(new Exception(""), "StartApplication " + application.Name, 3);
                    return Json(app.Start(), JsonRequestBehavior.AllowGet);
                }
            }
            catch (Exception exception)
            {
                Specific_Logging(exception, "StartApplication " + application.Name);
                response["status"] = "KO";
                response["details"] = "Exception " + exception.Message;
                return Json(response, JsonRequestBehavior.AllowGet);
            }
            response["status"] = "KO";
            response["details"] = "Une erreur inconnue est survenue lors de l'exécution de la commande ";
            Specific_Logging(new Exception(""), "StartApplication " + application.Name, 2);
            return Json(response, JsonRequestBehavior.AllowGet);
        }

        public JsonResult RestartApplication(int id)
        {
            Dictionary<string, string> response = new Dictionary<string, string>();
            response.Add("status", "");
            response.Add("details", "");
            Application application = db.Applications.Find(id);
            if (application == null)
            {
                response["status"] = "KO";
                response["details"] = "L'application n'a pas pu redémarrer car elle n'a pas été retrouvée dans la base de données.";
                return Json(response, JsonRequestBehavior.AllowGet);
            }
            try
            {
                ApplicationInfo app = new ApplicationInfo(application);
                Specific_Logging(new Exception(""), "RestartApplication " + application.Name, 3);
                return Json(app.Restart(), JsonRequestBehavior.AllowGet);
            }
            catch (Exception exception)
            {
                Specific_Logging(exception, "RestartApplication " + application.Name);
                response["status"] = "KO";
                response["details"] = "Exception " + exception.Message;
                return Json(response, JsonRequestBehavior.AllowGet);
            }
        }

        public JsonResult StopApplication(int id)
        {
            Dictionary<string, string> response = new Dictionary<string, string>();
            response.Add("status", "");
            response.Add("details", "");
            Application application = db.Applications.Find(id);
            if (application == null)
            {
                response["status"] = "KO";
                response["details"] = "L'application n'a pas été arrêtée car elle n'a pas été retrouvée dans la base de données.";
                return Json(response, JsonRequestBehavior.AllowGet);
            }
            try
            {
                ApplicationInfo app = new ApplicationInfo(application);
                Dictionary<string, string> status = app.GetState();
                if (status["status"] == "KO")
                {
                    response["status"] = "OK";
                    response["details"] = "L'application est déjà arrêtée.";
                }
                else
                {
                    Specific_Logging(new Exception(""), "StopApplication " + application.Name, 3);
                    return Json(app.Stop(), JsonRequestBehavior.AllowGet);
                }
            }
            catch (Exception exception)
            {
                Specific_Logging(exception, "StopApplication " + application.Name);
                response["status"] = "KO";
                response["details"] = "Exception " + exception.Message;
                return Json(response, JsonRequestBehavior.AllowGet);
            }
            Specific_Logging(new Exception(""), "StopApplication " + application.Name, 2);
            response["status"] = "KO";
            response["details"] = "Une erreur inconnue est survenue lors de l'exécution de la commande ";
            return Json(response, JsonRequestBehavior.AllowGet);
        }

        public string Export()
        {
            string message = "", FileName = "";
            MyApplication = new Excel.Application();
            MyApplication.Visible = false;
            try
            {
                List<AppDomain> domains = db.AppDomains.ToList();
                List<string> domains_list = new List<string>();
                if (domains.Count != 0)
                {
                    //Feed the database With The Domains
                    MyApplication = new Excel.Application();
                    MyApplication.Visible = false;

                    MyWorkbook = MyApplication.Workbooks.Open(HomeController.APP_DEFAULT_INIT_FILE_README);
                    MySheet = (Excel.Worksheet)MyWorkbook.Sheets.Add();
                    MySheet.Name = "Domains";
                    MySheet.Activate();

                    MySheet = (Excel.Worksheet)MyWorkbook.Sheets["Domains"];
                    int lastRow = MySheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row;
                    Excel.Range ActualRange = MySheet.get_Range("A" + lastRow,
                            "B" + lastRow);

                    MySheet.Cells[lastRow, 1] = "DOMAINES";
                    MySheet.Cells[lastRow, 1].EntireColumn.Font.Bold = true;

                    MySheet.Cells[lastRow, 1].EntireColumn.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                    MySheet.Cells[lastRow, 1].EntireColumn.ColumnWidth = 25;
                    MySheet.Cells[lastRow, 1].EntireColumn.Interior.Color = System.Drawing.ColorTranslator.FromHtml("#a5a5a5");
                    MySheet.Cells[lastRow, 1].Interior.Color = System.Drawing.ColorTranslator.FromHtml("#fff");

                    lastRow++;
                    foreach (AppDomain domain in domains.OrderBy(dom => dom.Name))
                    {
                        MySheet.Cells[lastRow, 1] = domain.Name;
                        domains_list.Add(domain.Name);
                        lastRow++;
                    }
                    //END DOMAINS FEEDING
                    MySheet = (Excel.Worksheet)MyWorkbook.Sheets.Add();
                    MySheet.Name = "Applications";
                    MySheet.Activate();

                    //APPLICATION FEEDING
                    List<Application> applications = db.Applications.ToList();
                    if (applications.Count != 0)
                    {
                        //Feed the database With The Apps
                        var validationDomainList = string.Join(",", domains_list.ToArray());
                        var navigatorList = string.Join(",", HomeController.APP_NAVIGATORS_LIST);

                        MySheet = (Excel.Worksheet)MyWorkbook.Sheets["Applications"];
                        lastRow = MySheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row;
                        ActualRange = MySheet.get_Range("A" + lastRow,
                                "Z" + lastRow);

                        MySheet.Cells[lastRow, 1] = "APPLICATIONS";
                        MySheet.Cells[lastRow, 1].EntireColumn.Font.Bold = true;
                        MySheet.Cells[lastRow, 2] = "SERVEURS";
                        MySheet.Cells[lastRow, 2].EntireColumn.Font.Bold = true;
                        MySheet.Cells[lastRow, 3] = "PROCEDURE DE DEMARRAGE (TYPE|CIBLE|ACTION;TYPE|CIBLE|ACTION;…)";
                        MySheet.Cells[lastRow, 3].EntireColumn.Font.Bold = true;
                        MySheet.Cells[lastRow, 4] = "PROCEDURE D'ARRET (TYPE|CIBLE|ACTION;TYPE|CIBLE|ACTION;…)";
                        MySheet.Cells[lastRow, 4].EntireColumn.Font.Bold = true;
                        MySheet.Cells[lastRow, 5] = "URL";
                        MySheet.Cells[lastRow, 5].EntireColumn.Font.Bold = true;
                        MySheet.Cells[lastRow, 6] = "ELEMENTS D'AUTHENTIFICATION";
                        MySheet.Cells[lastRow, 6].EntireColumn.Font.Bold = true;
                        MySheet.Cells[lastRow, 7] = "DOMAINES";
                        MySheet.Cells[lastRow, 7].EntireColumn.Font.Bold = true;
                        MySheet.Cells[lastRow, 8] = "NAVIGATEUR";
                        MySheet.Cells[lastRow, 8].EntireColumn.Font.Bold = true;

                        MySheet.Cells[lastRow, 1].EntireRow.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                        MySheet.Cells[lastRow, 1].EntireColumn.ColumnWidth = 25;
                        MySheet.Cells[lastRow, 1].EntireColumn.Interior.Color = System.Drawing.ColorTranslator.FromHtml("#93cddd");
                        MySheet.Cells[lastRow, 2].EntireColumn.ColumnWidth = 15;
                        MySheet.Cells[lastRow, 2].EntireColumn.Interior.Color = System.Drawing.ColorTranslator.FromHtml("#bfbfbf");
                        MySheet.Cells[lastRow, 3].EntireColumn.ColumnWidth = 35;
                        MySheet.Cells[lastRow, 3].EntireColumn.Interior.Color = System.Drawing.ColorTranslator.FromHtml("#c2d69a");
                        MySheet.Cells[lastRow, 4].EntireColumn.ColumnWidth = 35;
                        MySheet.Cells[lastRow, 4].EntireColumn.Interior.Color = System.Drawing.ColorTranslator.FromHtml("#d99795");
                        MySheet.Cells[lastRow, 5].EntireColumn.ColumnWidth = 35;
                        MySheet.Cells[lastRow, 5].EntireColumn.Interior.Color = System.Drawing.ColorTranslator.FromHtml("#8db4e3");
                        MySheet.Cells[lastRow, 6].EntireColumn.ColumnWidth = 35;
                        MySheet.Cells[lastRow, 6].EntireColumn.Interior.Color = System.Drawing.ColorTranslator.FromHtml("#fac090");
                        MySheet.Cells[lastRow, 7].EntireColumn.ColumnWidth = 35;
                        MySheet.Cells[lastRow, 7].EntireColumn.Interior.Color = System.Drawing.ColorTranslator.FromHtml("#a3a3a3");
                        MySheet.Cells[lastRow, 8].EntireColumn.ColumnWidth = 15;
                        MySheet.Cells[lastRow, 8].EntireColumn.Interior.Color = System.Drawing.ColorTranslator.FromHtml("#fff");

                        MySheet.Cells[lastRow, 1].EntireRow.Interior.Color = System.Drawing.ColorTranslator.FromHtml("#fff");
                        lastRow++;

                        foreach (Application application in applications.OrderBy(app => app.Name))
                        {
                            MySheet.Cells[lastRow, 1] = application.Name;
                            MySheet.Cells[lastRow, 5] = application.Url;
                            string elements = "";
                            foreach (AppHtmlElement element in application.AppHtmlElements)
                            {
                                elements += "Tagname=" + element.TagName +
                                    "|Xpath=" + element.AttrXpath +
                                    "|Id=" + element.AttrId + "|Name=" + element.AttrName +
                                    "|Class=" + element.AttrClass + "|Value=" + element.Value +
                                    "|Type=" + element.Type + ";";
                            }
                            if (elements.Length > 0)
                            {
                                elements = elements.Substring(0, elements.Length - 1);
                            }
                            MySheet.Cells[lastRow, 6] = elements;
                            MySheet.Cells[lastRow, 7] = db.AppDomains.Find(application.Domain).Name;
                            MySheet.Cells[lastRow, 8] = application.Navigator;

                            Excel.Range to_merge = (application.AppServers.Count == 0) ?
                                MySheet.get_Range("A" + lastRow, "A" + lastRow) :
                                MySheet.get_Range("A" + lastRow, "A" + (lastRow + application.AppServers.Count - 1));
                            to_merge.Merge();
                            to_merge.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                            to_merge.Font.Bold = true;
                            to_merge.Style.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                            to_merge.Style.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;


                            Excel.Range url_to_merge = (application.AppServers.Count == 0) ?
                                MySheet.get_Range("E" + lastRow, "E" + (lastRow + application.AppServers.Count)) :
                                MySheet.get_Range("E" + lastRow, "E" + (lastRow + application.AppServers.Count - 1));
                            url_to_merge.Merge();
                            url_to_merge.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                            url_to_merge.Font.Bold = true;
                            url_to_merge.Style.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                            url_to_merge.Style.HorizontalAlignment = Excel.XlHAlign.xlHAlignRight;

                            Excel.Range auth_to_merge = (application.AppServers.Count == 0) ?
                                MySheet.get_Range("F" + lastRow, "F" + (lastRow + application.AppServers.Count)) :
                                MySheet.get_Range("F" + lastRow, "F" + (lastRow + application.AppServers.Count - 1));
                            auth_to_merge.Merge();
                            auth_to_merge.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                            auth_to_merge.Font.Bold = true;
                            auth_to_merge.Style.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                            auth_to_merge.Style.HorizontalAlignment = Excel.XlHAlign.xlHAlignRight;

                            Excel.Range domain_to_merge = (application.AppServers.Count == 0) ?
                                MySheet.get_Range("G" + lastRow, "G" + (lastRow + application.AppServers.Count)) :
                                MySheet.get_Range("G" + lastRow, "G" + (lastRow + application.AppServers.Count - 1));
                            domain_to_merge.Merge();
                            domain_to_merge.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                            domain_to_merge.Font.Bold = true;
                            domain_to_merge.Style.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                            domain_to_merge.Style.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;


                            domain_to_merge.Validation.Delete();
                            domain_to_merge.Validation.Add(
                                Excel.XlDVType.xlValidateList,
                                Excel.XlDVAlertStyle.xlValidAlertInformation,
                                Excel.XlFormatConditionOperator.xlBetween,
                                validationDomainList,
                                Type.Missing);
                            domain_to_merge.Validation.IgnoreBlank = false;
                            domain_to_merge.Validation.InCellDropdown = true;

                            Excel.Range navigator_to_merge = (application.AppServers.Count == 0) ?
                                MySheet.get_Range("H" + lastRow, "H" + (lastRow + application.AppServers.Count)) :
                                MySheet.get_Range("H" + lastRow, "H" + (lastRow + application.AppServers.Count - 1));
                            navigator_to_merge.Merge();
                            navigator_to_merge.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                            navigator_to_merge.Font.Bold = true;
                            navigator_to_merge.Style.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                            navigator_to_merge.Style.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

                            navigator_to_merge.Validation.Add(
                                Excel.XlDVType.xlValidateList,
                                Excel.XlDVAlertStyle.xlValidAlertInformation,
                                Excel.XlFormatConditionOperator.xlBetween,
                                navigatorList,
                                Type.Missing);
                            navigator_to_merge.Validation.IgnoreBlank = false;
                            navigator_to_merge.Validation.InCellDropdown = true;

                            AppServer[] servers = application.AppServers.ToArray();
                            if (servers.Length != 0)
                            {
                                foreach (AppServer server in servers)
                                {
                                    MySheet.Cells[lastRow, 2] = server.Name;
                                    MySheet.Cells[lastRow, 2].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                                    MySheet.Cells[lastRow, 2].Font.Bold = true;
                                    MySheet.Cells[lastRow, 2].Style.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                                    MySheet.Cells[lastRow, 2].Style.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

                                    MySheet.Cells[lastRow, 3] = server.StartOrder;
                                    MySheet.Cells[lastRow, 3].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                                    MySheet.Cells[lastRow, 3].Font.Bold = true;
                                    MySheet.Cells[lastRow, 3].Style.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                                    MySheet.Cells[lastRow, 3].Style.HorizontalAlignment = Excel.XlHAlign.xlHAlignRight;

                                    MySheet.Cells[lastRow, 4] = server.StopOrder;
                                    MySheet.Cells[lastRow, 4].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                                    MySheet.Cells[lastRow, 4].Font.Bold = true;
                                    MySheet.Cells[lastRow, 4].Style.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                                    MySheet.Cells[lastRow, 4].Style.HorizontalAlignment = Excel.XlHAlign.xlHAlignRight;
                                    lastRow++;
                                }
                            }
                            else
                            {

                                MySheet.Cells[lastRow, 2].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                                MySheet.Cells[lastRow, 2].Font.Bold = true;
                                MySheet.Cells[lastRow, 2].Style.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                                MySheet.Cells[lastRow, 2].Style.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

                                MySheet.Cells[lastRow, 3].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                                MySheet.Cells[lastRow, 3].Font.Bold = true;
                                MySheet.Cells[lastRow, 3].Style.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                                MySheet.Cells[lastRow, 3].Style.HorizontalAlignment = Excel.XlHAlign.xlHAlignRight;

                                MySheet.Cells[lastRow, 4].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                                MySheet.Cells[lastRow, 4].Font.Bold = true;
                                MySheet.Cells[lastRow, 4].Style.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                                MySheet.Cells[lastRow, 4].Style.HorizontalAlignment = Excel.XlHAlign.xlHAlignRight;
                                lastRow++;
                            }
                        }
                        FileName = HomeController.APP_RESULTS_FOLDER + "ExportApps" + DateTime.Now.ToString("dd") +
                            DateTime.Now.ToString("MM") + DateTime.Now.ToString("yyyy") + ".xlsx";
                        try
                        {
                            if (System.IO.File.Exists(FileName))
                            {
                                System.IO.File.Delete(FileName);
                            }
                        }
                        catch { }

                        MyWorkbook.SaveAs(FileName,
                            Excel.XlFileFormat.xlOpenXMLWorkbook, null, null, null, null,
                            Excel.XlSaveAsAccessMode.xlNoChange, null, null, null, null, null);
                        message = "OK";
                    }
                    else
                    {
                        message = "La liste des applications de la base de données est vide.";
                    }
                }
                else
                {
                    message = "La liste des domaines de la base de données est vide.";
                }

            }
            catch (Exception exception)
            {
                Specific_Logging(exception, "Export");
                return "Erreur lors de l'ajout de l'application";
            }
            finally
            {
                McoUtilities.CloseExcel(MyApplication, MyWorkbook, MySheet);
            }
            if (message == "OK")
            {
                try
                {
                    System.Web.HttpResponse response = System.Web.HttpContext.Current.Response;
                    response.ClearContent();
                    response.Clear();
                    response.ContentType = "text/plain";
                    response.AddHeader("Content-Disposition", "attachment; filename=" + FileName + ";");
                    String RelativePath = FileName.Replace(Request.ServerVariables["APPL_PHYSICAL_PATH"], String.Empty);
                    response.TransmitFile(FileName);
                    response.Flush();
                    response.End();
                    return "OK";
                }
                catch (Exception exception)
                {
                    Specific_Logging(exception, "Export");
                    return exception.Message;
                }
            }
            Specific_Logging(new Exception(""), "Export", 2);
            return message;
        }

        [HttpPost]
        public string Import(bool import)
        {
            string message = "";
            if (import)
            {
                Purge();
                AppReport[] reports = db.AppReports.ToArray();
                foreach (AppReport report in reports)
                {
                    message += DeleteAppReport(report.Id) + "\n <br />";
                }

                Application[] applications = db.Applications.ToArray();
                foreach (Application application in applications)
                {
                    message += DeleteApplication(application.Id) + "\n <br />";
                }

                AppDomain[] domains = db.AppDomains.ToArray();
                foreach (AppDomain domain in domains)
                {
                    message += DeleteAppDomain(domain.Id) + "\n <br />";
                }

                //Feed the database With The Apps
                MyApplication = new Excel.Application();
                MyApplication.Visible = false;
                try
                {
                    Dictionary<int, ServersController.VirtualizedServer> FOREST = ServersController.GetInformationsFromForestDomains();
                    ReftechServers[] REFTECH_SERVERS = null;
                    try
                    {
                        REFTECH_SERVERS = db.ReftechServers.ToArray();
                    }
                    catch { }
                    MyWorkbook = MyApplication.Workbooks.Open(HomeController.APP_INIT_FILE);
                    //DOMAIN FEEDING
                    MySheet = (Excel.Worksheet)MyWorkbook.Sheets["Domains"];
                    int lastRow = MySheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row;
                    for (int index = 2; index <= lastRow; index++)
                    {
                        try
                        {
                            System.Array MyValues = (System.Array)MySheet.get_Range("A" +
                                    index.ToString(), "B" + index.ToString()).Cells.Value;
                            if (MyValues.GetValue(1, 1) == null)
                            {
                                continue;
                            }
                            string DomainName = MyValues.GetValue(1, 1).ToString().Trim();
                            if (DomainName != null && DomainName.Trim() != "" &&
                                db.AppDomains.Where(dom => dom.Name == DomainName).Count() == 0)
                            {
                                AppDomain domain = db.AppDomains.Create();
                                domain.Name = DomainName;
                                domain.Applications = 0;
                                if (ModelState.IsValid)
                                {
                                    db.AppDomains.Add(domain);
                                    db.SaveChanges();
                                }
                            }
                            else
                            {
                                continue;
                            }
                        }
                        catch
                        {
                            continue;
                        }
                    }
                    //END FEEDING DOMAINS


                    //FEEDING APPLICATIONS
                    MySheet = (Excel.Worksheet)MyWorkbook.Sheets["Applications"];
                    lastRow = MySheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row;
                    string ApplicationName = "";
                    for (int index = 2; index <= lastRow; index++)
                    {
                        try
                        {
                            System.Array MyValues = (System.Array)MySheet.get_Range("A" +
                                index.ToString(), "H" + index.ToString()).Cells.Value;

                            if (MyValues.GetValue(1, 1) == null && MyValues.GetValue(1, 2) == null)
                            {
                                ApplicationName = "";
                                continue;
                            }
                            Application application = null;
                            ApplicationName = (MyValues.GetValue(1, 1) == null) ? ApplicationName :
                                MyValues.GetValue(1, 1).ToString().Trim();
                            if (db.Applications.Where(app => app.Name == ApplicationName).Count() == 1)
                            {
                                application = db.Applications.Where(app => app.Name == ApplicationName).First();
                                if (application.Url == "")
                                {
                                    string url = (MyValues.GetValue(1, 5) != null) ?
                                        MyValues.GetValue(1, 5).ToString().Trim() : "";
                                    if (url != "" && (!url.ToLower().Contains("http://") && !url.ToLower().Contains("https://")))
                                    {
                                        url = "http://" + url;
                                    }
                                    application.Url = url;
                                    if (ModelState.IsValid)
                                    {
                                        db.Entry(application).State = System.Data.Entity.EntityState.Modified;
                                        db.SaveChanges();
                                        message += "Une Url a été rajoutée avec succès aux informations de l'application " + application.Name + ". \n <br />";
                                    }
                                    else
                                    {
                                        message += "Erreur lors de la mise à jour de l'application " + ApplicationName + ". \n <br />";
                                    }
                                }
                                if (application.AppHtmlElements.Count == 0)
                                {
                                    //AUTH PROCEDURES
                                    if (MyValues.GetValue(1, 6) != null && MyValues.GetValue(1, 6).ToString().Trim() != "")
                                    {
                                        string[] elements = MyValues.GetValue(1, 6).ToString().Split(';');
                                        foreach (string element in elements)
                                        {
                                            if (element.Trim() != "")
                                            {
                                                AppHtmlElement node = db.AppHtmlElements.Create();
                                                node.ApplicationId = application.Id;
                                                string[] infos = element.Split('|');
                                                foreach (string info in infos)
                                                {
                                                    if (info.Trim() != "" && info.Contains("Tagname"))
                                                    {
                                                        node.TagName = info.Substring(info.IndexOf("=") + 1);
                                                    }
                                                    if (info.Trim() != "" && info.Contains("Xpath"))
                                                    {
                                                        node.AttrXpath = info.Substring(info.IndexOf("=") + 1);
                                                    }
                                                    if (info.Trim() != "" && info.Contains("Id"))
                                                    {
                                                        node.AttrId = info.Substring(info.IndexOf("=") + 1);
                                                    }
                                                    if (info.Trim() != "" && info.Contains("Name"))
                                                    {
                                                        node.AttrName = info.Substring(info.IndexOf("=") + 1);
                                                    }
                                                    if (info.Trim() != "" && info.Contains("Class"))
                                                    {
                                                        node.AttrClass = info.Substring(info.IndexOf("=") + 1);
                                                    }
                                                    if (info.Trim() != "" && info.Contains("Value"))
                                                    {
                                                        node.Value = info.Substring(info.IndexOf("=") + 1);
                                                    }
                                                    if (info.Trim() != "" && info.Contains("Type"))
                                                    {
                                                        node.Type = info.Substring(info.IndexOf("=") + 1);
                                                    }
                                                }
                                                if (ModelState.IsValid)
                                                {
                                                    db.AppHtmlElements.Add(node);
                                                    db.SaveChanges();
                                                }
                                            }
                                        }
                                    }

                                }
                            }
                            else
                            {
                                application = db.Applications.Create();
                                application.Name = ApplicationName;
                                string domain = (MyValues.GetValue(1, 7) != null) ?
                                    MyValues.GetValue(1, 7).ToString().Trim() : "";
                                if (domain != "")
                                {
                                    AppDomain appdomain = db.AppDomains.Where(dom => dom.Name == domain).FirstOrDefault();
                                    if (appdomain != null)
                                    {
                                        application.Domain = appdomain.Id;
                                    }
                                    else
                                    {
                                        continue;
                                    }
                                }
                                else
                                {
                                    continue;
                                }
                                string navigator = (MyValues.GetValue(1, 8) != null) ?
                                    MyValues.GetValue(1, 8).ToString().Trim() : "IE";
                                application.Navigator = navigator;

                                string url = (MyValues.GetValue(1, 5) != null) ?
                                    MyValues.GetValue(1, 5).ToString().Trim() : "";
                                if (url != "" && (!url.ToLower().Contains("http://") && !url.ToLower().Contains("https://")))
                                {
                                    url = "http://" + url;
                                    application.Url = url;
                                }
                                else
                                {
                                    if (url != "" && (url.ToLower().Contains("http://") || url.ToLower().Contains("https://")))
                                    {
                                        application.Url = url;
                                    }
                                }
                                if (ModelState.IsValid)
                                {
                                    db.Applications.Add(application);
                                    db.SaveChanges();
                                    AppDomain appdomain = db.AppDomains.Find(application.Domain);
                                    if (appdomain != null)
                                    {
                                        appdomain.Applications = db.Applications.Where(app => app.Domain == appdomain.Id).Count();
                                        db.Entry(appdomain).State = System.Data.Entity.EntityState.Modified;
                                        db.SaveChanges();
                                    }
                                    message += "---------------------------------------------------------------\n <br />";
                                    message += "L'application " + application.Name + " a été ajoutée avec succès. \n <br />";
                                    //AUTH PROCEDURES
                                    if (MyValues.GetValue(1, 6) != null && MyValues.GetValue(1, 6).ToString().Trim() != "")
                                    {
                                        string[] elements = MyValues.GetValue(1, 6).ToString().Split(';');
                                        foreach (string element in elements)
                                        {
                                            if (element.Trim() != "")
                                            {
                                                AppHtmlElement node = db.AppHtmlElements.Create();
                                                node.ApplicationId = application.Id;
                                                string[] infos = element.Split('|');
                                                foreach (string info in infos)
                                                {
                                                    if (info.Trim() != "" && info.Contains("Tagname"))
                                                    {
                                                        node.TagName = info.Substring(info.IndexOf("=") + 1);
                                                    }
                                                    if (info.Trim() != "" && info.Contains("Xpath"))
                                                    {
                                                        node.AttrXpath = info.Substring(info.IndexOf("=") + 1);
                                                    }
                                                    if (info.Trim() != "" && info.Contains("Id"))
                                                    {
                                                        node.AttrId = info.Substring(info.IndexOf("=") + 1);
                                                    }
                                                    if (info.Trim() != "" && info.Contains("Name"))
                                                    {
                                                        node.AttrName = info.Substring(info.IndexOf("=") + 1);
                                                    }
                                                    if (info.Trim() != "" && info.Contains("Class"))
                                                    {
                                                        node.AttrClass = info.Substring(info.IndexOf("=") + 1);
                                                    }
                                                    if (info.Trim() != "" && info.Contains("Value"))
                                                    {
                                                        node.Value = info.Substring(info.IndexOf("=") + 1);
                                                    }
                                                    if (info.Trim() != "" && info.Contains("Type"))
                                                    {
                                                        node.Type = info.Substring(info.IndexOf("=") + 1);
                                                    }
                                                }
                                                if (ModelState.IsValid)
                                                {
                                                    db.AppHtmlElements.Add(node);
                                                    db.SaveChanges();
                                                }
                                            }
                                        }
                                    }
                                }
                                else
                                {
                                    message += "---------------------------------------------------------------\n <br />";
                                    message += "Erreur lors de l'ajout de " + ApplicationName + ". \n <br />";
                                    message += "---------------------------------------------------------------\n <br />";
                                    continue;
                                }
                            }
                            if (MyValues.GetValue(1, 2) != null && MyValues.GetValue(1, 2).ToString().Trim() != "")
                            {
                                AppServer server = db.AppServers.Create();
                                server.Name = MyValues.GetValue(1, 2).ToString().Trim().ToUpper();
                                ServersController.VirtualizedServer_Result virtual_server = ServersController.GetServerInformations(FOREST, REFTECH_SERVERS, server, HomeController.APP_MODULE, false);
                                server.StartOrder = server.StopOrder = "";
                                server.ApplicationId = application.Id;

                                string[] Procedures = MyValues.GetValue(1, 3).ToString().Split(';');
                                string[] ReverseProcedures = Procedures.Reverse().ToArray();
                                bool default_stop_procedures = true;
                                if (MyValues.GetValue(1, 4) != null && MyValues.GetValue(1, 4).ToString().Trim() != "")
                                {
                                    ReverseProcedures = MyValues.GetValue(1, 4).ToString().Split(';');
                                    default_stop_procedures = false;
                                }

                                //StartOrder
                                foreach (string Procedure in Procedures)
                                {
                                    if (Procedure.Trim() != "")
                                    {
                                        string[] Details = Procedure.Split('|');
                                        string type = Details[0].Trim();
                                        string target = Details[1].Trim();
                                        string action = Details[2].Trim();
                                        server.StartOrder += type + "|" + target + "|" + action + ";";
                                    }
                                }
                                if (server.StartOrder.Length > 0)
                                {
                                    server.StartOrder = server.StartOrder.Substring(0, server.StartOrder.Length - 1);
                                }
                                //StopOrder
                                foreach (string Procedure in ReverseProcedures)
                                {
                                    if (Procedure.Trim() != "")
                                    {
                                        string[] Details = Procedure.Split('|');
                                        string type = Details[0].Trim();
                                        string target = Details[1].Trim();
                                        string action = Details[2].Trim();
                                        if (default_stop_procedures)
                                        {
                                            if (type == "BATCH")
                                            {
                                                continue;
                                            }
                                            switch (action)
                                            {
                                                case "START":
                                                    server.StopOrder += type + "|" + target + "|STOP;";
                                                    break;
                                                case "RESTART":
                                                    server.StopOrder += type + "|" + target + "|STOP;";
                                                    break;
                                                case "CHECK":
                                                    continue;
                                            }
                                        }
                                        else
                                        {
                                            server.StopOrder += type + "|" + target + "|" + action + ";";
                                        }

                                    }
                                }
                                if (server.StopOrder.Length > 0)
                                {
                                    server.StopOrder = server.StopOrder.Substring(0, server.StopOrder.Length - 1);
                                }
                                if (ModelState.IsValid)
                                {
                                    db.AppServers.Add(server);
                                    db.SaveChanges();
                                    message += "Le serveur " + server.Name + " a été ajoutée avec succès à l'application " + application.Name + ". \n <br />";
                                }
                                else
                                {
                                    message += "Le serveur " + server.Name + " n'a pas pu être rajouté à l'application " + application.Name + ". \n <br />";
                                }
                            }


                        }
                        catch (Exception damned)
                        {
                            Specific_Logging(damned, "Import");
                            message += "Exception " + damned.Message;
                            continue;
                        }
                    }
                }
                catch (Exception exception)
                {
                    Specific_Logging(exception, "Import");
                    return "Erreur lors de l'ajout de l'application";
                }
                finally
                {
                    McoUtilities.CloseExcel(MyApplication, MyWorkbook, MySheet);
                }
                Specific_Logging(new Exception(""), "Import", 3);
                return message;
            }
            else
            {
                Specific_Logging(new Exception(""), "Import", 2);
                return "Erreur d'importation.";
            }
        }

        public string DownloadInitFile()
        {
            try
            {
                System.Web.HttpResponse response = System.Web.HttpContext.Current.Response;
                response.ClearContent();
                response.Clear();
                response.ContentType = "text/plain";
                string filepath = HomeController.APP_RELATIVE_INIT_FILE;
                response.AddHeader("Content-Disposition", "attachment; filename=" + filepath + ";");
                String RelativePath = HomeController.APP_DEFAULT_INIT_FILE.Replace(Request.ServerVariables["APPL_PHYSICAL_PATH"], String.Empty);
                response.TransmitFile(HomeController.APP_DEFAULT_INIT_FILE);
                response.Flush();
                response.End();
                return "OK";
            }
            catch (Exception exception)
            {
                Specific_Logging(exception, "DownloadInitFile");
                return exception.Message;
            }
        }

        public Dictionary<Application, BrowseUrlResult> BrowseApplications(List<Application> applications)
        {
            Dictionary<Application, BrowseUrlResult> BrowseResults = new Dictionary<Application, BrowseUrlResult>();
            List<ApplicationInfo> list = new List<ApplicationInfo>();
            Dictionary<ApplicationInfo, BrowseUrlResult> VirtualBrowseResults = new Dictionary<ApplicationInfo, BrowseUrlResult>();
            foreach (Application application in applications)
            {
                BrowseResults.Add(application, new BrowseUrlResult(application));
                list.Add(new ApplicationInfo(application));
            }

            VirtualBrowseResults = ApplicationInfo.BrowseApplicationInfos(list);

            foreach (ApplicationInfo app in list)
            {
                Application application = app.GetApplication(BrowseResults);
                if (application != null)
                {
                    BrowseResults[application] = VirtualBrowseResults[app];
                }
            }
            return BrowseResults;
        }

        public Dictionary<Application, BrowseUrlResult> BrowseApplication(int id)
        {
            Dictionary<Application, BrowseUrlResult> BrowseResults = new Dictionary<Application, BrowseUrlResult>();
            List<ApplicationInfo> list = new List<ApplicationInfo>();
            Dictionary<ApplicationInfo, BrowseUrlResult> VirtualBrowseResults = new Dictionary<ApplicationInfo, BrowseUrlResult>();
            Application application = db.Applications.Find(id);
            if (application == null)
            {
                return null;

            }
            BrowseResults.Add(application, new BrowseUrlResult(application));
            list.Add(new ApplicationInfo(application));
            VirtualBrowseResults = ApplicationInfo.BrowseApplicationInfos(list);

            foreach (ApplicationInfo app in list)
            {
                application = app.GetApplication(BrowseResults);
                if (application != null)
                {
                    BrowseResults[application] = VirtualBrowseResults[app];
                }
            }
            return BrowseResults;
        }

        public Dictionary<AppServer, List<BrowseUrlResult>> BrowseServers(List<AppServer> servers)
        {
            Dictionary<AppServer, List<BrowseUrlResult>> BrowseResults = new Dictionary<AppServer, List<BrowseUrlResult>>();
            List<AppServerInfo> list = new List<AppServerInfo>();
            Dictionary<AppServerInfo, List<BrowseUrlResult>> VirtualBrowseResults = new Dictionary<AppServerInfo, List<BrowseUrlResult>>();
            foreach (AppServer server in servers)
            {
                BrowseResults.Add(server, new List<BrowseUrlResult>());
                list.Add(new AppServerInfo(server, server.Application));
            }

            VirtualBrowseResults = AppServerInfo.BrowseServerInfos(list);

            foreach (AppServerInfo serv in list)
            {
                AppServer server = serv.GetServer(BrowseResults);
                if (server != null)
                {
                    BrowseResults[server] = VirtualBrowseResults[serv];
                }
            }
            return BrowseResults;
        }

        public JsonResult GetNavigatorsList()
        {
            string[] Navigators = HomeController.APP_NAVIGATORS_LIST;
            return Json(Navigators, JsonRequestBehavior.AllowGet);
        }

        public JsonResult GetProcedureTypes()
        {
            string[] Types = HomeController.APP_PROCEDURE_TYPES;
            return Json(Types, JsonRequestBehavior.AllowGet);
        }

        public JsonResult GetProcedureActions()
        {
            string[] Actions = HomeController.APP_PROCEDURE_ACTIONS;
            return Json(Actions, JsonRequestBehavior.AllowGet);
        }

        public JsonResult GetApplicationServers(int id)
        {
            Application application = db.Applications.Find(id);
            if (application == null)
            {
                return null;
            }
            return Json(application.AppServers.ToArray(), JsonRequestBehavior.AllowGet);
        }

        public JsonResult ExecuteSchedule(int id)
        {
            if (!CanCheck())
            {
                return NotifyImpossibility(true);
            }
            AppSchedule schedule = db.AppSchedules.Find(id);
            if (schedule == null)
            {
                Specific_Logging(new Exception(""), "ExecuteSchedule", 1);
                return null;
            }
            else
            {
                schedule.State = "En cours";
                if (ModelState.IsValid)
                {
                    db.Entry(schedule).State = System.Data.Entity.EntityState.Modified;
                    db.SaveChanges();
                }
            }
            Dictionary<string, string> results = new Dictionary<string, string>();
            results["response"] = "OK";
            results["email"] = "";
            results["errors"] = "";
            results["applications"] = "";

            List<Application> SelectedApps = new List<Application>();
            foreach (Scheduled_Application app in schedule.Scheduled_Applications)
            {
                SelectedApps.Add(app.Application);
            }
            List<AppDomain> domains = new List<AppDomain>();
            foreach (Application application in SelectedApps)
            {
                AppDomain domain = db.AppDomains.Find(application.Domain);
                if (!domains.Contains(domain))
                {
                    domains.Add(domain);
                }
            }
            if (SelectedApps.Count == 0)
            {
                results["response"] = "Aucune Application n'a été sélectionnée dans la base de données";
                results["email"] = "";
                results["errors"] = "";
                return Json(results, JsonRequestBehavior.AllowGet);
            }

            int emailId = 0;
            string ExecutionErrors = "";
            try
            {
                MyApplication = new Excel.Application();
                MyApplication.Visible = false;
                MyWorkbook = MyApplication.Workbooks.Add(Excel.XlWBATemplate.xlWBATWorksheet);
                MySheet = (Excel.Worksheet)MyWorkbook.Sheets.Add();
                MySheet.Name = "Check" + DateTime.Now.ToString("dd") +
                    DateTime.Now.ToString("MM") + DateTime.Now.ToString("yyyy");
                MySheet.Activate();

                AppReport report = db.AppReports.Create();
                report.DateTime = DateTime.Now;
                report.TotalChecked = 0;
                report.TotalErrors = 0;
                report.ResultPath = "";
                report.Module = HomeController.APP_MODULE;
                report.ScheduleId = schedule.Id;
                report.Schedule = schedule;
                report.Author = HomeController.SYSTEM_IDENTITY;
                Email email = db.Emails.Create();
                report.Email = email;
                email.Report = report;
                email.Module = HomeController.APP_MODULE;
                email.Recipients = "";
                email = Emails_Controller.SetRecipients(email, HomeController.APP_MODULE);
                if (ModelState.IsValid)
                {
                    db.AppReports.Add(report);
                    db.SaveChanges();
                    emailId = report.Email.Id;
                    int reportNumber = db.AppReports.Count();
                    if (reportNumber > HomeController.APP_MAX_REPORT_NUMBER)
                    {
                        int reportNumberToDelete = reportNumber - HomeController.APP_MAX_REPORT_NUMBER;
                        AppReport[] reportsToDelete =
                            db.AppReports.OrderBy(ide => ide.Id).Take(reportNumberToDelete).ToArray();
                        foreach (AppReport toDeleteReport in reportsToDelete)
                        {
                            DeleteAppReport(toDeleteReport.Id);
                        }
                    }
                }
                else
                {
                    results["response"] = "KO";
                    results["email"] = null;
                    results["errors"] = "Impossible de créer un rapport dans la base de données.";
                }

                //START OF TREATMENT
                int lastRow = MySheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row;
                SelectedApps = SelectedApps.OrderBy(app => app.Name).ToList();

                foreach (AppDomain domain in domains.OrderBy(dom => dom.Name))
                {
                    List<Application> dom_applications = SelectedApps.Where(app => app.Domain == domain.Id).ToList();
                    MySheet.Cells[lastRow, 1] = domain.Name;
                    int firstline = lastRow;
                    int mergedlines = 0;
                    foreach (Application application in dom_applications.OrderBy(app => app.Name))
                    {
                        Application_Report applicationReport = new Application_Report();
                        applicationReport.Application = application;
                        applicationReport.State = "KO";
                        applicationReport.Details = "";
                        applicationReport.Authentified = "";
                        applicationReport.Linkable = "";
                        applicationReport.AppReportId = report.Id;
                        applicationReport.ApplicationId = application.Id;
                        applicationReport.AppReport = report;
                        if (ModelState.IsValid)
                        {
                            try
                            {
                                db.ApplicationReports.Add(applicationReport);
                                db.SaveChanges();
                            }
                            catch { }
                        }
                        report.TotalChecked++;
                        if (applicationReport.Authentified.Trim() == "")
                        {
                            List<Application> pplication = new List<Application>();
                            pplication.Add(application);

                            Dictionary<Application, BrowseUrlResult> BrowseResults
                                = new Dictionary<Application, BrowseUrlResult>();
                            BrowseResults = BrowseApplications(pplication);

                            applicationReport.Authentified = (BrowseResults[application].Status == "OK") ? "OK" : "KO : " + BrowseResults[application].Details;

                        }
                        if (applicationReport.Authentified != "")
                        {
                            MySheet.Cells[lastRow, 6] = "Authentification : " + applicationReport.Authentified;
                        }
                        else
                        {
                            MySheet.Cells[lastRow, 6] = "Authentification : Inconnue";
                        }

                        AppServer[] servers = application.AppServers.ToArray();
                        Excel.Range ActualRange = MySheet.get_Range("A" + lastRow,
                                "Z" + lastRow);

                        int lines = servers.Length;
                        lines = (lines == 0) ? 1 : lines;
                        mergedlines += lines;

                        if (servers.Length == 0)
                        {
                            MySheet.Cells[lastRow, 2] = application.Name;
                            MySheet.Cells[lastRow, 2].EntireColumn.ColumnWidth = 25;
                            MySheet.Cells[lastRow, 2].EntireRow.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                            MySheet.Cells[lastRow, 2].EntireRow.Font.Color = System.Drawing.ColorTranslator.FromHtml("#000");
                            MySheet.Cells[lastRow, 2].EntireRow.Style.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                            MySheet.Cells[lastRow, 2].EntireRow.Style.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                            MySheet.Cells[lastRow, 2].EntireRow.Font.Bold = true;
                            MySheet.Cells[lastRow, 2].EntireRow.Interior.Color = System.Drawing.ColorTranslator.FromHtml("#eeece1");
                            MySheet.Hyperlinks.Add(MySheet.Cells[lastRow, 2], application.Url, Type.Missing, application.Name, application.Name);

                            MySheet.Cells[lastRow, 6].Interior.Color = System.Drawing.ColorTranslator.FromHtml("#fff");
                            MySheet.Cells[lastRow, 6].EntireColumn.ColumnWidth = 40;
                            MySheet.Cells[lastRow, 6].Font.Color = (applicationReport.Authentified == "OK") ? System.Drawing.ColorTranslator.FromHtml("#22b14c")
                                   : System.Drawing.ColorTranslator.FromHtml("#ff2f00");

                            MySheet.Cells[lastRow, 3].EntireColumn.ColumnWidth = 15;
                            MySheet.Cells[lastRow, 4].EntireColumn.ColumnWidth = 10;
                            MySheet.Cells[lastRow, 5].EntireColumn.ColumnWidth = 10;
                            MySheet.Cells[lastRow, 5].Font.Color = System.Drawing.ColorTranslator.FromHtml("#de5a26");

                            applicationReport.State = "";


                            if (applicationReport.Authentified != "")
                            {
                                MySheet.Cells[lastRow, 6] = "Authentification : " + applicationReport.Authentified;
                            }
                            else
                            {
                                MySheet.Cells[lastRow, 6] = "Authentification : Inconnue";
                            }
                            if (ModelState.IsValid)
                            {
                                db.Entry(applicationReport).State = System.Data.Entity.EntityState.Modified;
                                db.SaveChanges();
                            }
                            lastRow += 1;
                        }
                        else
                        {
                            MySheet.Cells[lastRow, 2] = application.Name;
                            MySheet.Cells[lastRow, 2].EntireColumn.ColumnWidth = 25;
                            MySheet.Cells[lastRow, 6].Font.Color = (applicationReport.Authentified == "OK") ? System.Drawing.ColorTranslator.FromHtml("#22b14c")
                                   : System.Drawing.ColorTranslator.FromHtml("#ff2f00");

                            Excel.Range to_merge = MySheet.get_Range("B" + lastRow, "B" + (lastRow + application.AppServers.Count - 1));
                            to_merge.Merge();
                            to_merge.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                            to_merge.Font.Bold = true;
                            to_merge.Style.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                            to_merge.Style.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

                            Excel.Range auth_to_merge = MySheet.get_Range("F" + lastRow, "F" + (lastRow + application.AppServers.Count - 1));
                            auth_to_merge.Merge();
                            auth_to_merge.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                            auth_to_merge.Font.Bold = true;
                            auth_to_merge.Style.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                            auth_to_merge.Style.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

                            List<string> Status = new List<string>();
                            foreach (AppServer server in servers)
                            {
                                MySheet.Cells[lastRow, 3] = server.Name;
                                MySheet.Cells[lastRow, 3].EntireRow.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                                MySheet.Cells[lastRow, 3].EntireRow.Font.Color = System.Drawing.ColorTranslator.FromHtml("#000");
                                MySheet.Cells[lastRow, 3].EntireRow.Style.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                                MySheet.Cells[lastRow, 3].EntireRow.Style.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                                MySheet.Cells[lastRow, 3].EntireRow.Font.Bold = true;

                                MySheet.Cells[lastRow, 3].EntireColumn.ColumnWidth = 15;



                                JsonResult JsonResponse = GetAppServerState(server.Id);
                                Dictionary<string, string> response = (Dictionary<string, string>)JsonResponse.Data;

                                AppServer_Report serverreport = db.AppServerReports.Create();
                                serverreport.AppServer = server;
                                serverreport.Application_Report = applicationReport;
                                serverreport.State = serverreport.Details = serverreport.Ping = "";// response["status"];
                                serverreport.State = response["status"];


                                if (serverreport.State != "OK")
                                {
                                    Status.Add("KO");
                                    try
                                    {
                                        Ping ping = new Ping();
                                        PingOptions options = new PingOptions(64, true);
                                        PingReply pingreply = ping.Send(server.Name);
                                        serverreport.Ping = (pingreply.Status.ToString() == "Success") ? "OK" : "KO";
                                    }
                                    catch
                                    {
                                        serverreport.Ping = "KO";
                                    }
                                    MySheet.Cells[lastRow, 3].EntireRow.Interior.Color = System.Drawing.ColorTranslator.FromHtml("#ff3f3f");
                                    MySheet.Cells[lastRow, 5].Interior.Color = System.Drawing.ColorTranslator.FromHtml("#fff");
                                    MySheet.Cells[lastRow, 5].EntireColumn.ColumnWidth = 10;
                                    MySheet.Cells[lastRow, 4] = "KO";
                                    MySheet.Cells[lastRow, 5] = "Ping: " + serverreport.Ping;

                                    MySheet.Cells[lastRow, 5].Font.Color = (serverreport.Ping == "OK") ? System.Drawing.ColorTranslator.FromHtml("#22b14c")
                                        : System.Drawing.ColorTranslator.FromHtml("#ff2f00");
                                    serverreport.Details = "";
                                    int column = 7;
                                    string[] Details = response["details"].Split(new string[] { "\n" }, StringSplitOptions.None);
                                    foreach (string infos in Details)
                                    {
                                        if (infos.Trim() != "")
                                        {
                                            serverreport.Details += infos + " | ";
                                            MySheet.Cells[lastRow, column] = infos;
                                            MySheet.Cells[lastRow, column].EntireColumn.ColumnWidth = 35;
                                            column++;
                                        }
                                    }
                                    if (serverreport.Details.Length > 1)
                                    {
                                        serverreport.Details = serverreport.Details.Substring(0, serverreport.Details.Length - 3);
                                    }
                                }
                                else
                                {
                                    Status.Add("OK");
                                    serverreport.Ping = "OK";
                                    MySheet.Cells[lastRow, 3].EntireRow.Interior.Color = System.Drawing.ColorTranslator.FromHtml("#22b14c");
                                    MySheet.Cells[lastRow, 5].Interior.Color = System.Drawing.ColorTranslator.FromHtml("#fff");
                                    MySheet.Cells[lastRow, 5].EntireColumn.ColumnWidth = 10;
                                    MySheet.Cells[lastRow, 4] = "OK";
                                    MySheet.Cells[lastRow, 5] = "Ping: OK";
                                    MySheet.Cells[lastRow, 5].Font.Color = System.Drawing.ColorTranslator.FromHtml("#22b14c");
                                    serverreport.Details = "";
                                }

                                if (ModelState.IsValid)
                                {
                                    db.AppServerReports.Add(serverreport);
                                    db.SaveChanges();
                                }

                                auth_to_merge.Style.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                                auth_to_merge.Interior.Color = System.Drawing.ColorTranslator.FromHtml("#fff");
                                auth_to_merge.EntireColumn.ColumnWidth = 40;

                                auth_to_merge.Font.Color = (applicationReport.Authentified == "OK") ? System.Drawing.ColorTranslator.FromHtml("#22b14c")
                                           : System.Drawing.ColorTranslator.FromHtml("#ff2f00");
                                auth_to_merge.EntireColumn.ColumnWidth = 40;
                                if (applicationReport.Authentified != "")
                                {
                                    MySheet.Cells[lastRow, 6] = "Authentification : " + applicationReport.Authentified;
                                }
                                else
                                {
                                    MySheet.Cells[lastRow, 6] = "Authentification : Inconnue";

                                }
                                lastRow += 1;
                            }
                            to_merge.Hyperlinks.Add(to_merge, application.Url, Type.Missing, application.Name, application.Name);
                            if (!Status.Contains("KO"))
                            {
                                applicationReport.State = "OK";
                            }
                            else
                            {
                                report.TotalErrors++;
                                if (!Status.Contains("OK"))
                                {
                                    applicationReport.State = "KO";
                                }
                                else
                                {
                                    applicationReport.State = "H-OK";
                                }
                            }
                            if (ModelState.IsValid)
                            {
                                db.Entry(applicationReport).State = System.Data.Entity.EntityState.Modified;
                                db.SaveChanges();
                            }
                            string backgroundcolor = (applicationReport.State == "OK") ? "#22b14c" : (applicationReport.State == "KO") ? "#ff3f3f" :
                                (applicationReport.State == "H-OK") ? "#de5a26" : "#eeece1";
                            to_merge.Interior.Color = System.Drawing.ColorTranslator.FromHtml(backgroundcolor);
                        }
                    }
                    mergedlines = (mergedlines == 0) ? 1 : mergedlines;
                    Excel.Range dom_to_merge = MySheet.get_Range("A" + firstline, "A" + (firstline + mergedlines - 1));
                    dom_to_merge.Merge();
                    dom_to_merge.Interior.Color = System.Drawing.ColorTranslator.FromHtml("#dcdbdb");
                    dom_to_merge.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                    dom_to_merge.Font.Bold = true;
                    dom_to_merge.EntireColumn.ColumnWidth = 35;
                    dom_to_merge.Style.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                    dom_to_merge.Style.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                }
                string FileName = HomeController.APP_RESULTS_FOLDER + "Check Applications " + DateTime.Now.ToString("dd") +
                    DateTime.Now.ToString("MM") + DateTime.Now.ToString("yyyy") + " - " + report.Id + ".xlsx";
                MyWorkbook.SaveAs(FileName,
                    Excel.XlFileFormat.xlOpenXMLWorkbook, null, null, null, null,
                    Excel.XlSaveAsAccessMode.xlNoChange, null, null, null, null, null);
                report.Duration = DateTime.Now.Subtract(report.DateTime);
                report.ResultPath = FileName;
                if (ModelState.IsValid)
                {
                    db.Entry(report).State = System.Data.Entity.EntityState.Modified;
                    db.SaveChanges();
                    string buildOk = BuildEmail(email.Id);
                    if (buildOk != "BuildOK")
                    {
                        ExecutionErrors += "Erreur lors de la mise à jour du mail \n <br />";
                    }
                }
                else
                {
                    results["response"] = "KO";
                    results["email"] = null;
                    results["errors"] = "Echec lors de l'enregistrement dans la base de données.";
                }
            }
            catch (Exception exception)
            {
                Specific_Logging(exception, "ExecuteSchedule");
            }
            finally
            {
                McoUtilities.CloseExcel(MyApplication, MyWorkbook, MySheet);
            }

            results["email"] = emailId.ToString();
            results["errors"] = "Fin d'exécution. \n" + "Erreurs d'exécution : " + ExecutionErrors;
            Emails_Controller.AutoSend(emailId);
            schedule.State = (schedule.Multiplicity != "Une fois") ? "Planifié" : "Terminé";
            schedule.NextExecution = Schedules_Controller.GetNextExecution(schedule);
            schedule.Executed++;
            if (ModelState.IsValid)
            {
                db.Entry(schedule).State = System.Data.Entity.EntityState.Modified;
                db.SaveChanges();
            }
            Specific_Logging(new Exception(""), "ExecuteSchedule", 3);
            return Json(results, JsonRequestBehavior.AllowGet);
        }

        public JsonResult NotifyImpossibility(bool autosend = true)
        {
            return McoUtilities.NotifyImpossibility(HomeController.APP_MODULE, autosend);
        }

        public bool CanCheck()
        {
            if (!McoUtilities.CheckIfLoggedOn(HomeController.DEFAULT_DOMAIN_AND_USERNAME_IMPERSONNATION))
            {
                return false;
            }
            return true;
        }

        public string HasUnachieviedReport()
        {
            return Reports_Controller.HasUnachieviedReport(HomeController.APP_MODULE);
        }

        private void Specific_Logging(Exception exception, string action, int level = 0)
        {
            string author = "UNKNOWN";
            if (User != null && User.Identity != null && User.Identity.Name != " ")
            {
                author = User.Identity.Name;
            }
            McoUtilities.Specific_Logging(exception, action, HomeController.APP_MODULE, level, author);
        }
    }
}
