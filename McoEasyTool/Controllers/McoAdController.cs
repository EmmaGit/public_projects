using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Security.Principal;
using System.Text;
using System.Web.Mvc;
using DotNet.Highcharts;
using DotNet.Highcharts.Enums;
using DotNet.Highcharts.Helpers;
using DotNet.Highcharts.Options;

namespace McoEasyTool.Controllers
{
    [AllowAnonymous]
    public class McoAdController : Controller
    {
        private DataModelContainer db = new DataModelContainer();
        private ReportsController Reports_Controller = new ReportsController();
        private EmailsController Emails_Controller = new EmailsController();
        private SchedulesController Schedules_Controller = new SchedulesController();
        private ServersController Servers_Controller = new ServersController();
        private McoUtilities.UNCAccessWithCredentials UNC_ACCESSOR = new McoUtilities.UNCAccessWithCredentials();

        public ActionResult Home()
        {
            ViewBag.AD_DESC_0 = McoUtilities.GetModuleDescription(HomeController.AD_MODULE, 0);
            ViewBag.AD_DESC_1 = McoUtilities.GetModuleDescription(HomeController.AD_MODULE, 1);
            ViewBag.AD_DESC_2 = McoUtilities.GetModuleDescription(HomeController.AD_MODULE, 2);
            ViewBag.AD_DESC_3 = McoUtilities.GetModuleDescription(HomeController.AD_MODULE, 3);
            ViewBag.AD_DESC_4 = McoUtilities.GetModuleDescription(HomeController.AD_MODULE, 4);
            return View();
        }

        public ActionResult DisplaySchedules()
        {
            return View(db.AdSchedules.OrderBy(name => name.TaskName).ToList());
        }

        public ActionResult DisplaySettings()
        {
            AD_Settings ad_settings = LoadSettings();
            if (ad_settings == null)
            {
                return HttpNotFound();
            }
            ViewBag.Message = "Paramétrage du filtrage Check AD";
            return View(ad_settings);
        }

        public ActionResult DisplayReports()
        {
            return View(db.AdReports.OrderByDescending(report => report.Id).ToList());
        }

        public ActionResult DisplayScheduleReports(int id)
        {
            AdSchedule schedule = db.AdSchedules.Find(id);
            object[] boundaries =
                McoUtilities.GetIdValues<AdSchedule>(schedule, HomeController.OBJECT_ATTR_ID);
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
            ICollection<AdReport> reports = db.AdReports.Where(report => report.ScheduleId == id).ToList();
            return View(reports.OrderByDescending(report => report.Id).ToList());
        }

        public ActionResult DisplayFaultyServers()
        {
            List<FaultyServer> faultyservers = db.FaultyServers.ToList();
            foreach (FaultyServer faultyserver in faultyservers)
            {
                if (faultyserver.FaultyServer_Reports.Count == 0)
                {
                    db.FaultyServers.Remove(faultyserver);
                    db.SaveChanges();
                }
            }
            return View(db.FaultyServers.ToList());
        }

        public ActionResult DisplayRecipients()
        {
            return View(db.Recipients.Where(rec => rec.Module == HomeController.AD_MODULE).ToList());
        }

        public ActionResult DisplayFaultyServerStatistics()
        {
            FaultyServer[] faultyservers = db.FaultyServers.ToArray();
            String[] servernames = new String[faultyservers.Length];
            Object[] errorserver = new Object[faultyservers.Length];
            int index = 0;
            foreach (FaultyServer faultyserver in faultyservers)
            {
                servernames[index] = faultyserver.Name;
                errorserver[index] = faultyserver.FaultyServer_Reports.Count;
                index++;
            }
            string start = "";
            string stop = "";
            if (db.AdReports.Count() != 0)
            {
                start = db.AdReports.Min(date => date.DateTime).ToString();
                stop = db.AdReports.Max(date => date.DateTime).ToString();
            }
            else
            {
                start = DateTime.Now.ToString();
                stop = DateTime.Now.AddHours(1).ToString();
            }

            Highcharts chart = new Highcharts("chart")
                .InitChart(new Chart
                {
                    DefaultSeriesType = ChartTypes.Column,
                    MarginRight = 130,
                    MarginBottom = 25,
                    BackgroundColor = new BackColorOrGradient(ColorTranslator.FromHtml("#ded9d4")),
                    ClassName = "FaultyServer"
                })
                .SetTitle(new Title
                {
                    Text = "Statistiques des défaillances",
                    X = -20
                })
                .SetSubtitle(new Subtitle
                {
                    Text = "Le nombre de rapport ayant référencé ces contrôleurs de domaine entre le " + start + " et le " + stop,
                    X = -20
                })
                .SetXAxis(new XAxis
                {
                    Categories = servernames
                })
                .SetYAxis(new YAxis
                {
                    Title = new YAxisTitle { Text = "Nombre de rapports" },
                    PlotLines = new[]
                            {
                                new YAxisPlotLines
                                    {
                                        Value = 0,
                                        Width = 1,
                                        Color = ColorTranslator.FromHtml("#808080")
                                    }
                            }
                })
                .SetTooltip(new Tooltip
                {
                    Formatter = @"function() {
                                        return '<b>'+ this.series.name +'</b><br/>'+
                                    this.x +': '+ this.y +' références';
                                }",
                    BackgroundColor = new BackColorOrGradient(ColorTranslator.FromHtml("#fff")),
                    BorderColor = ColorTranslator.FromHtml("transparent"),
                })
                .SetLegend(new Legend
                {
                    Layout = Layouts.Vertical,
                    Align = HorizontalAligns.Right,
                    VerticalAlign = VerticalAligns.Top,
                    X = -10,
                    Y = 100,
                    BorderWidth = 0
                })
                .SetSeries(new[]
                    {
                        new Series { Name = "Références", Data = new Data(errorserver),Color = ColorTranslator.FromHtml("#ff3f3f") },
                    }
                );
            return View(chart);
        }

        public ActionResult DisplayFaultyServerDetails(int id)
        {
            FaultyServer faultyserver = db.FaultyServers.Find(id);
            if (faultyserver == null)
            {
                return HttpNotFound();
            }
            object[] boundaries =
                McoUtilities.GetIdValues<FaultyServer>(faultyserver, HomeController.OBJECT_ATTR_ID);
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
            return View(faultyserver);
        }

        public ActionResult DisplayScheduleVersionFaultyServerDetails(int id)
        {
            FaultyServer faultyserver = db.FaultyServers.Find(id);
            if (faultyserver == null)
            {
                return HttpNotFound();
            }
            return View(faultyserver);
        }

        public ActionResult DisplayReportVersionFaultyServerDetails(int id)
        {
            FaultyServer faultyserver = db.FaultyServers.Find(id);
            if (faultyserver == null)
            {
                return HttpNotFound();
            }
            return View(faultyserver);
        }

        public ActionResult DisplayReportDetails(int id)
        {
            AdReport report = db.AdReports.Find(id);
            if (report == null)
            {
                return HttpNotFound();
            }
            object[] boundaries =
                McoUtilities.GetIdValues<AdReport>(report, HomeController.OBJECT_ATTR_ID, true);
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
            return View(db.FaultyServerReports.Where(rep => rep.AdReportId == id).ToList());
        }

        public ActionResult DisplayReportStatistics()
        {
            AdReport[] reports = db.AdReports.ToArray();
            String[] reportsDate = new String[reports.Length];
            Object[] totalerrors = new Object[reports.Length];
            Object[] fatalerrors = new Object[reports.Length];
            int index = 0;
            foreach (AdReport report in reports)
            {
                reportsDate[index] = report.DateTime.ToString();
                totalerrors[index] = report.TotalErrors;
                fatalerrors[index] = report.FatalErrors;
                index++;
            }

            System.Drawing.Color[] ColumnColors = { };

            Highcharts chart = new Highcharts("chart")
                .InitChart(new Chart
                {
                    DefaultSeriesType = ChartTypes.Column,
                    MarginRight = 130,
                    MarginBottom = 25,
                    BorderWidth = 0,
                    BorderRadius = 15,
                    PlotBackgroundColor = null,
                    PlotShadow = false,
                    PlotBorderWidth = 0,
                    BackgroundColor = new BackColorOrGradient(ColorTranslator.FromHtml("#ded9d4")),
                    ClassName = "Report"
                })
                .SetOptions(new GlobalOptions
                {
                    Colors = new[]
                                         {
                                             ColorTranslator.FromHtml("#DDDF0D"),
                                             ColorTranslator.FromHtml("#7798BF"),
                                             ColorTranslator.FromHtml("#55BF3B"),
                                             ColorTranslator.FromHtml("#DF5353"),
                                             ColorTranslator.FromHtml("#DDDF0D"),
                                             ColorTranslator.FromHtml("#aaeeee"),
                                             ColorTranslator.FromHtml("#ff0066"),
                                             ColorTranslator.FromHtml("#eeaaee")

                                         }
                })
                .SetTitle(new Title
                {
                    Text = "Contrôleurs défaillant",
                    X = -20
                })
                .SetSubtitle(new Subtitle
                {
                    Text = "Le nombre de contrôleurs défaillant ayant été détectés ces derniers jours",
                    X = -20
                })
                .SetXAxis(new XAxis
                {
                    Categories = reportsDate
                })
                .SetYAxis(new YAxis
                {
                    Title = new YAxisTitle { Text = "Nombre de contrôleurs" },
                    PlotLines = new[]
                            {
                                new YAxisPlotLines
                                    {
                                        Value = 0,
                                        Width = 1,
                                        Color = ColorTranslator.FromHtml("#f00")
                                    }
                            }
                })
                .SetTooltip(new Tooltip
                {
                    Formatter = @"function() {
                                        return '<b>'+ this.series.name +'</b><br/>'+
                                    this.x +': '+ this.y +' contrôleurs défaillant';
                                }",
                    BorderWidth = 0,
                    Shadow = false,
                    BackgroundColor = new BackColorOrGradient(ColorTranslator.FromHtml("#fff")),
                    BorderColor = ColorTranslator.FromHtml("transparent"),
                    //Enabled = false

                })
                .SetLegend(new Legend
                {
                    Layout = Layouts.Vertical,
                    Align = HorizontalAligns.Right,
                    VerticalAlign = VerticalAligns.Top,
                    X = -10,
                    Y = 100,
                    BorderWidth = 0
                })
                .SetSeries(new[]
                    {
                        new Series { Name = "Serveurs défaillant", Data = new Data(totalerrors),Color = ColorTranslator.FromHtml("#e75114") },
                        new Series { Name = "Serveurs en défaillance critique", Data = new Data(fatalerrors),Color = ColorTranslator.FromHtml("#ff3f3f") },
                    }
                );

            return View(chart);

        }

        public ActionResult DisplayScheduleReportDetails(int id)
        {
            AdReport report = db.AdReports.Find(id);
            AdSchedule schedule = db.AdSchedules.Find(report.ScheduleId);
            object[] boundaries =
                McoUtilities.GetIdValues<AdReport>(report, HomeController.OBJECT_ATTR_ID, true);
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
            return View(db.FaultyServerReports.Where(rep => rep.AdReportId == id).ToList());
        }

        public string CreateSchedule()
        {
            int day = 0; int month = 0; int year = 0; int hours = 0; int minutes = 0;
            string autocorrect = "false";
            string postedtaskname = Request.Form["taskname"].ToString();
            bool dayOk = Int32.TryParse(Request.Form["day"].ToString(), out day);
            bool monthOk = Int32.TryParse(Request.Form["month"].ToString(), out month);
            bool yearOk = Int32.TryParse(Request.Form["year"].ToString(), out year);
            bool hoursOk = Int32.TryParse(Request.Form["hours"].ToString(), out hours);
            bool minutesOk = Int32.TryParse(Request.Form["minutes"].ToString(), out minutes);
            string multiplicity = Request.Form["multiplicity"].ToString();
            autocorrect = Request.Form["autocorrect"].ToString();

            string result = "";
            if (dayOk && monthOk && yearOk && hoursOk && minutesOk)
            {
                TimeSpan time = new TimeSpan(hours, minutes, 30);
                DateTime now = DateTime.Now;
                DateTime scheduled = new DateTime(year, month + 1, day) + time;
                int scheduleId = 0;
                if (scheduled.CompareTo(DateTime.Now) > 0)
                {
                    string taskname = HomeController.AD_MODULE + " AutoCheck ";

                    AdSchedule schedule = db.AdSchedules.Create();
                    schedule.CreationTime = DateTime.Now;
                    schedule.NextExecution = scheduled;
                    schedule.Generator = User.Identity.Name;
                    schedule.Multiplicity = multiplicity;
                    schedule.Executed = 0;
                    schedule.State = "Planifié";
                    schedule.Module = HomeController.AD_MODULE;
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
                        if (autocorrect == "true")
                        {
                            schedule.AutoCorrect = true;
                        }
                        else
                        {
                            schedule.AutoCorrect = false;
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
            AdSchedule schedule = db.AdSchedules.Find(id);
            if (schedule == null)
            {
                return HttpNotFound().ToString();
            }

            int day = 0; int month = 0; int year = 0; int hours = 0; int minutes = 0;
            string autocorrect = "false";
            string postedtaskname = Request.Form["taskname"].ToString();
            bool dayOk = Int32.TryParse(Request.Form["day"].ToString(), out day);
            bool monthOk = Int32.TryParse(Request.Form["month"].ToString(), out month);
            bool yearOk = Int32.TryParse(Request.Form["year"].ToString(), out year);
            bool hoursOk = Int32.TryParse(Request.Form["hours"].ToString(), out hours);
            bool minutesOk = Int32.TryParse(Request.Form["minutes"].ToString(), out minutes);
            string multiplicity = Request.Form["multiplicity"].ToString();
            autocorrect = Request.Form["autocorrect"].ToString();
            string result = "";
            if (dayOk && monthOk && yearOk && hoursOk && minutesOk)
            {
                TimeSpan time = new TimeSpan(hours, minutes, 30);
                DateTime now = DateTime.Now;
                DateTime scheduled = new DateTime(year, month + 1, day) + time;
                if (scheduled.CompareTo(DateTime.Now) > 0 &&
                    (Schedules_Controller.Delete(schedule) == "La tâche a été correctement supprimée"))
                {
                    string taskname = HomeController.AD_MODULE + " AutoCheck ";

                    schedule.NextExecution = scheduled;
                    schedule.Generator = User.Identity.Name;
                    schedule.Multiplicity = multiplicity;
                    schedule.State = "Planifié";
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
                        if (autocorrect == "true")
                        {
                            schedule.AutoCorrect = true;
                        }
                        else
                        {
                            schedule.AutoCorrect = false;
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
            AdSchedule schedule = db.AdSchedules.Find(id);
            if (schedule.Reports.Count != 0)
            {
                List<Report> reports = schedule.Reports.ToList();
                foreach (AdReport report in reports)
                {
                    DeleteAdReport(report.Id);
                    db.Entry(schedule).State = System.Data.Entity.EntityState.Modified;
                    db.SaveChanges();
                }
            }
            string result = Schedules_Controller.Delete(schedule);
            schedule = db.AdSchedules.Find(id);
            if (result == "La tâche a été correctement supprimée")
            {
                db.AdSchedules.Remove(schedule);
                db.SaveChanges();
            }
            Specific_Logging(new Exception(""), "DeleteSchedule", 3);
            return result;
        }

        public string ReSendLastEmail(int id)
        {
            AdSchedule schedule = db.AdSchedules.Find(id);
            if (schedule != null)
            {
                if (db.Reports.Where(report => report.ScheduleId == id).Count() != 0)
                {
                    Report report = db.Reports.Where(rep => rep.ScheduleId == id).OrderByDescending(rep => rep.Id).First();
                    return Reports_Controller.ReSend(report.Id);
                }
                return "Cette tâche planifiée n'a pour l'instant généré aucun rapport, ou alors ils ont été supprimés.";

            }
            Specific_Logging(new Exception(""), "ResendLastEmail", 3);
            return "Cette tâche planifiée n'a pas été trouvée dans la base de données.";
        }

        public AD_Settings LoadSettings()
        {
            AD_Settings ad_settings;
            if (db.AD_Settings.Count() == 0)
            {
                ad_settings = new AD_Settings();
                ad_settings.DurationFilter = true;
                ad_settings.ErrorFilter = false;
                ad_settings.PingFilter = true;
                ad_settings.StateFilter = false;
                ad_settings.Duration = 2;
                ad_settings.MemoryErrors = "8606";
                ad_settings.State = "Operationnel";
                if (ModelState.IsValid)
                {
                    db.AD_Settings.Add(ad_settings);
                    db.SaveChanges();
                }
            }
            else
            {
                ad_settings = db.AD_Settings.First();
            }
            return ad_settings;
        }

        [HttpPost]
        public int SaveSettings()
        {
            AD_Settings ad_settings = LoadSettings();
            try
            {
                string test = Request.Form["DurationFilter"].ToString();
                ad_settings.DurationFilter = (Request.Form["DurationFilter"].ToString().ToLower() == "true");
                ad_settings.PingFilter = (Request.Form["PingFilter"].ToString().ToLower() == "true");
                ad_settings.StateFilter = (Request.Form["StateFilter"].ToString().ToLower() == "true");
                ad_settings.ErrorFilter = (Request.Form["ErrorFilter"].ToString().ToLower() == "true");
                ad_settings.MemoryErrors = Request.Form["MemoryErrors"].ToString();
                if (ad_settings.DurationFilter)
                {
                    int duration = 1;
                    Int32.TryParse(Request.Form["Duration"].ToString(), out duration);
                    ad_settings.Duration = duration;
                }
                if (ad_settings.StateFilter)
                {
                    ad_settings.State = Request.Form["State"].ToString();
                }
                if (ad_settings.ErrorFilter)
                {
                    ad_settings.Errors = Request.Form["Errors"].ToString();
                }
                if (ModelState.IsValid)
                {
                    db.Entry(ad_settings).State = System.Data.Entity.EntityState.Modified;
                    db.SaveChanges();
                    Specific_Logging(new Exception("...."), "SaveSettings", 3);
                    return 1;
                }
            }
            catch (Exception exception)
            {
                Specific_Logging(exception, "SaveSettings");
                return 0;
            }
            Specific_Logging(new Exception("...."), "SaveSettings", 2);
            return 0;
        }

        public string ViewEmails(int id)
        {
            AdReport report = db.AdReports.Find(id);
            if (report != null)
            {
                return Reports_Controller.ViewEmail(report.Id);
            }
            return "Ce rapport n'a pas été retrouvé dans la base de données.";
        }

        public string ReSendEmail(int id)
        {
            AdReport report = db.AdReports.Find(id);
            if (report != null)
            {
                return Reports_Controller.ReSend(report.Id);
            }
            return "Le rapport n'a pas été retrouvé dans la base de données.";
        }

        public string DownloadReport(int id)
        {
            AdReport report = db.AdReports.Find(id);
            if (report == null)
            {
                return HttpNotFound().ToString();
            }
            return Reports_Controller.Download(report.Id);
        }

        public string DeleteAdReport(int id)
        {
            try
            {
                AdReport report = db.AdReports.Find(id);
                Email email = report.Email;

                List<FaultyServer_Report> faultyserverreports = db.FaultyServerReports.ToList();
                foreach (FaultyServer_Report faultyserverreport in faultyserverreports)
                {
                    if (faultyserverreport.AdReportId == report.Id)
                    {
                        db.FaultyServerReports.Remove(faultyserverreport);
                        db.SaveChanges();
                    }
                }

                List<FaultyServer> faultyservers = db.FaultyServers.ToList();
                foreach (FaultyServer faultyserver in faultyservers)
                {
                    if (faultyserver.FaultyServer_Reports.Count == 0)
                    {
                        DeleteFaultyServer(faultyserver.Id);
                    }
                }

                Specific_Logging(new Exception("...."), "DeleteAdReport", 3);
                return Reports_Controller.Delete(report.Id);
            }
            catch (Exception exception)
            {
                Specific_Logging(exception, "DeleteAdReport");
                return "Une erreur est surveunue lors de la suppression" +
                    exception.Message;
            }
        }

        public bool IsCritical(VirtualizedAdServer serverinfo, AD_Settings ad_settings)
        {
            bool critical = false;
            if (ad_settings.DurationFilter && serverinfo.Absence.IndexOf('d') != -1)
            {
                string adays = serverinfo.Absence.Substring(0, serverinfo.Absence.IndexOf('d'));
                int days = 0;
                bool isNum = Int32.TryParse(adays, out days);
                if (isNum)
                {
                    if (days >= ad_settings.Duration)
                    {
                        critical = true;
                    }
                }
            }
            if (ad_settings.PingFilter && serverinfo.Server_Result.Ping != "OK")
            {
                critical = true;
            }
            if (ad_settings.ErrorFilter)
            {
                string[] errors = ad_settings.Errors.Split(',');
                foreach (string error in errors)
                {
                    if (serverinfo.Errors.IndexOf(error) != -1)
                    {
                        critical = true;
                        break;
                    }
                }
            }
            if (ad_settings.StateFilter && ad_settings.State.Trim().ToLower() != serverinfo.Server_Result.AD_Server.Status.Trim().ToLower()
                && serverinfo.Server_Result.AD_Server.Status.Trim().ToLower() != "" && serverinfo.Server_Result.AD_Server.Status.Trim().ToLower() != "inconnu")
            {
                critical = false;
            }
            return critical;
        }

        public JsonResult Filter(JsonResult Jsonlaunch, int reportId)
        {
            AD_Settings ad_settings = LoadSettings();
            string Options = "<br />Ce rapport a été généré avec le paramétrage suivant:<br />";
            //DURATION
            Options += "Filtrage Durée d'absence : " + ((ad_settings.DurationFilter) ? "<span style='color:#0197bc;'>Oui</span>" : "Non");
            if (ad_settings.DurationFilter)
            {
                Options += " <span style='color:#0197bc;'> : " + ad_settings.Duration.ToString() + " jour(s)</span>";
            }
            Options += "<br />";
            //PING
            Options += "Filtrage Ping KO : " + ((ad_settings.PingFilter) ? "<span style='color:#0197bc;'>Oui</span>" : "Non");
            Options += "<br />";
            //STATE
            Options += "Filtrage Etat du contrôleur : " + ((ad_settings.StateFilter) ? "<span style='color:#0197bc;'>Oui</span>" : "Non");
            if (ad_settings.StateFilter)
            {
                Options += " Sélection des contrôleurs d'état <span style='color:#0197bc;'>" + ad_settings.State + "</span>";
            }
            Options += "<br />";
            //ERRORS
            Options += "Filtrage Codes d'erreurs : " + ((ad_settings.ErrorFilter) ? "<span style='color:#0197bc;'>Oui</span>" : "Non");
            if (ad_settings.ErrorFilter)
            {
                Options += " Les erreurs de code <span style='color:#0197bc;'>" + ad_settings.Errors + "</span> sont déclarées critiques.";
            }
            Options += "<br />";

            Dictionary<string, string[]> data = (Dictionary<string, string[]>)Jsonlaunch.Data;
            string[] content = data["content"];
            AdReport report = db.AdReports.Find(reportId);
            Email email = report.Email;
            string errors = "";
            bool startsplit = false;
            Dictionary<int, ServersController.VirtualizedServer> FOREST = ServersController.GetInformationsFromForestDomains();
            ReftechServers[] REFTECH_SERVERS = null;
            try
            {
                REFTECH_SERVERS = db.ReftechServers.ToArray();
            }
            catch { }
            Dictionary<int, VirtualizedAdServer> serverinfos = new Dictionary<int, VirtualizedAdServer>();
            Dictionary<int, VirtualizedAdServer> noserverinfos = new Dictionary<int, VirtualizedAdServer>();
            Dictionary<string, string> serverlist = new Dictionary<string, string>();
            int serverinfosIndex = 0;
            int noserverinfosIndex = 0;
            string origin = "Source DSA";

            Dictionary<string, string> results = new Dictionary<string, string>();
            //TREATMENT OF THE RESULT FILE
            int number = 0;
            foreach (string line in content)
            {
                number++;
                //Conditions to test if the parsing must start or not
                if (startsplit && (
                        (line.IndexOf("Les erreurs fonctionnelles suivantes", StringComparison.OrdinalIgnoreCase) != -1) ||
                        (line.IndexOf("Experienced the following operational errors", StringComparison.OrdinalIgnoreCase) != -1)
                   ))
                {
                    break;
                }
                if ((!startsplit) && (line.IndexOf("source", StringComparison.OrdinalIgnoreCase) != -1))
                {
                    if (line.IndexOf("dsa", StringComparison.OrdinalIgnoreCase) >= 0)
                    {
                        startsplit = true;
                        continue;
                    }
                }
                else
                {
                    if (!startsplit)
                    {
                        continue;
                    }
                }
                //When the parsing have started
                if (startsplit)
                {
                    if (line.Trim() == "")
                    {
                        continue;
                    }
                    if ((line.IndexOf("destination", StringComparison.OrdinalIgnoreCase) != -1)
                        && (line.IndexOf("dsa", StringComparison.OrdinalIgnoreCase) != -1))
                    {
                        origin = "Destination DSA";
                        continue;
                    }
                    int fail = 0; string servername = "", serverorigin = "", servererrors = "", serverabsence = "";
                    string[] seeker = line.Trim().Split(' ');
                    string[] failures = line.Substring(0, line.IndexOf("/")).Trim().Split(' ');
                    if (failures.Length > 0 && Int32.TryParse(failures[failures.Length - 1], out fail))
                    {
                        if (fail == 0)
                        {
                            continue;
                        }
                    }
                    if (seeker[0].Trim() != "")
                    {
                        servername = seeker[0];
                    }
                    serverorigin = origin;
                    if (line.IndexOf("(") - 1 > 0)
                    {
                        servererrors = line.Trim().Substring(line.IndexOf("(") - 1);
                    }
                    else
                    {
                        servererrors = "Empty";
                    }


                    //GET ABSENCE DURATION & FAILS FROM FILE LINE
                    if (seeker.Length < 7 && origin == "Destination DSA")
                    {
                        break;
                    }

                    for (int index = 1; index < seeker.Length; index++)
                    {

                        if (seeker[index].Trim() != "" && (serverabsence == null || serverabsence.Trim() == ""))
                        {
                            serverabsence = seeker[index];
                            continue;
                        }
                    }

                    //THERE IS MISTAKE, WE CAN LOOK FOR FURTHER INFOS
                    if (fail != 0)
                    {
                        VirtualizedAdServer virtual_server = new VirtualizedAdServer();
                        Server server = new Server();
                        server.Name = servername;
                        virtual_server.Server_Result = ServersController.GetServerInformations(FOREST, REFTECH_SERVERS, server, HomeController.AD_MODULE, true);
                        virtual_server.Fail = fail; virtual_server.Absence = serverabsence;
                        virtual_server.Origin = serverorigin; virtual_server.Errors = servererrors;

                        //DECIDE WEITHER THE FAILURES ARE CRITICAL OR NOT
                        if (IsCritical(virtual_server, ad_settings))
                        {
                            if (serverlist.ContainsKey(virtual_server.Server_Result.AD_Server.Name))
                            {
                                int location = 0;
                                Int32.TryParse(serverlist[servername].Split('-')[1], out location);
                                if (serverlist[servername].Split('-')[0] == "criticals")
                                {
                                    serverinfos[location] = (GetMaxDuration(serverinfos[location].Absence, virtual_server.Absence) != serverinfos[location].Absence) ?
                                                virtual_server : serverinfos[location];
                                }
                                else
                                {
                                    serverinfos.Add(serverinfosIndex, virtual_server);
                                    serverlist[servername] = "criticals-" + serverinfosIndex.ToString();
                                    serverinfosIndex++;

                                    //CLEAN NO SERVER INFO
                                    for (int index = 0; index < noserverinfos.Count - 1; index++)
                                    {
                                        if (noserverinfos[index].Server_Result.AD_Server.Name == virtual_server.Server_Result.AD_Server.Name)
                                        {
                                            location = index;
                                            noserverinfos.Remove(index);
                                            break;
                                        }
                                    }
                                    for (int index = location; index < noserverinfos.Count - 1; index++)
                                    {
                                        noserverinfos[index] = noserverinfos[index + 1];
                                        serverlist[noserverinfos[index].Server_Result.AD_Server.Name] = "nocriticals-" + index.ToString();
                                    }
                                    noserverinfos.Remove(noserverinfos.Count-1);
                                    noserverinfosIndex--;
                                }
                            }
                            else
                            {
                                serverinfos.Add(serverinfosIndex, virtual_server);
                                serverlist.Add(servername, "criticals-" + serverinfosIndex.ToString());
                                serverinfosIndex++;
                            }
                        }
                        else
                        {
                            if (serverlist.ContainsKey(servername))
                            {

                                int location = 0;
                                Int32.TryParse(serverlist[servername].Split('-')[1], out location);
                                if (serverlist[servername].Split('-')[0] == "nocriticals")
                                {
                                    noserverinfos[location] = (GetMaxDuration(noserverinfos[location].Absence, virtual_server.Absence) != noserverinfos[location].Absence) ?
                                                virtual_server : noserverinfos[location];
                                }
                            }
                            else
                            {
                                noserverinfos.Add(noserverinfosIndex, virtual_server);
                                serverlist.Add(servername, "nocriticals-" + noserverinfosIndex.ToString());
                                noserverinfosIndex++;
                            }
                        }
                    }//END OF FAULTY SERVERS TREATMENT
                }
            }//END OF TREATMENT OF THE RESULT FILE

            //CONSTRUCTION OF RESULTS
            if (serverinfos.Count() == 0 && noserverinfos.Count() == 0)
            {
                errors = "Une erreur s'est peut être produite car aucun serveur n'a été référencé.";
                ViewBag.Message = errors;
            }
            int fatalErrors = serverinfos.Count();
            int totalErrors = serverinfos.Count() + noserverinfos.Count();
            //All the faulty servers are supposed to be in the dictionnary
            FaultyServer[] faultyservers = null;
            if (db.FaultyServers.Count() > 0)
            {
                faultyservers = db.FaultyServers.ToArray();
            }

            string body = "<br/><style>tr:hover>td{cursor:pointer;background-color:#68b3ff;}</style><span style='color:#e75114;'>L'absence de tableaux indique qu'aucun serveur n'a présenté de défauts de réplication; prière " +
                    "de consulter le fichier Txt envoyé en pièce jointe pour plus de détails.</span><br /><br />";
            int serverId = 0;

            if (serverinfos.Count() > 0)
            {
                body += "Le tableau rouge contient les contrôleurs se trouvant dans un état critique de réplication. <br />";
                body += "<table style='position:relative;width:100%;'><thead>" +
                "<tr style='position:relative;width:100%;height:35px;background-color:#ff3f3f;'><th>Serveur</th><th>Durée d'absence</th><th>Ping</th>" +
                "<th>Etat</th><th>Origine</th><th>Site</th><th>IdSite</th><th>Erreurs</th><th>Détails</th></tr></thead><tbody>";
                for (int index = 0; index < serverinfos.Count(); index++)
                {
                    bool founded = false;
                    FaultyServer_Report faultyserverreport = db.FaultyServerReports.Create();
                    faultyserverreport.AdReportId = reportId;
                    faultyserverreport.AdReport = report;
                    faultyserverreport.AbsenceDuration = serverinfos[index].Absence;
                    faultyserverreport.Details = serverinfos[index].Errors;
                    faultyserverreport.Ping = serverinfos[index].Server_Result.Ping;
                    if (faultyservers != null && faultyservers.Length > 0)
                    {
                        foreach (FaultyServer server in faultyservers)
                        {
                            if (server.Name.Trim().ToString() == serverinfos[index].Server_Result.AD_Server.Name.Trim())
                            {
                                founded = true;
                                serverId = server.Id;
                                server.Status = serverinfos[index].Server_Result.AD_Server.Status;
                                server.Location = serverinfos[index].Server_Result.AD_Server.Location;
                                server.Site = serverinfos[index].Server_Result.AD_Server.Site;
                                server.ActiveDirecotryDomain = serverinfos[index].Server_Result.AD_Server.ActiveDirecotryDomain;
                                if ((server.IpAddress == null || server.IpAddress.Trim() == "") &&
                                    (serverinfos[index].Server_Result.AD_Server.IpAddress != null))
                                {
                                    server.IpAddress = serverinfos[index].Server_Result.AD_Server.IpAddress;
                                }
                                faultyserverreport.FaultyServer = server;
                                break;
                            }
                        }
                    }
                    if (!founded)
                    {
                        FaultyServer NewServer = db.FaultyServers.Create();
                        NewServer = serverinfos[index].Server_Result.AD_Server;
                        if (ModelState.IsValid)
                        {
                            db.FaultyServers.Add(NewServer);
                            db.SaveChanges();
                            serverId = NewServer.Id;
                            faultyserverreport.FaultyServer = NewServer;
                        }
                    }
                    
                    faultyserverreport.FaultyServerId = serverId;
                    
                    
                    
                    try
                    {
                        if (ModelState.IsValid)
                        {
                            db.FaultyServerReports.Add(faultyserverreport);
                            db.SaveChanges();
                        }
                    }
                    catch (Exception ce)
                    {
                        int a = 2;
                    }

                    string codeerror = serverinfos[index].Errors.Split('(')[1];
                    codeerror = codeerror.Split(')')[0];
                    body += "<tr style='position:relative;width:100%;min-height:35px;border:1px solid #000;text-align:center;'>" +
                        "<td style='position:relative;padding-left:1px;width:8%;border:1px solid #000;'>" + faultyserverreport.FaultyServer.Name + "</td>" +
                        "<td style='position:relative;padding-left:1px;width:11%;border:1px solid #000;'>" + faultyserverreport.AbsenceDuration + "</td>" +
                        "<td style='position:relative;padding-left:1px;width:5%;border:1px solid #000;'>" + faultyserverreport.Ping + "</td>" +
                        "<td style='position:relative;padding-left:1px;width:6%;border:1px solid #000;'>" + faultyserverreport.FaultyServer.Status + "</td>" +
                        "<td style='position:relative;padding-left:1px;width:5%;border:1px solid #000;'>" + serverinfos[index].Origin + "</td>" +
                        "<td style='position:relative;padding-left:1px;width:15%;border:1px solid #000;'>" + faultyserverreport.FaultyServer.Site + "</td>" +
                        "<td style='position:relative;padding-left:1px;width:20%;border:1px solid #000;'>" + faultyserverreport.FaultyServer.IdSite + "</td>" +
                        "<td style='position:relative;padding-left:1px;width:5%;border:1px solid #000;'>" + codeerror + "</td>" +
                        "<td style='position:relative;padding-left:1px;width:25%;border:1px solid #000;'>" + serverinfos[index].Errors.Substring(codeerror.Length + 2) + "</td></tr>";
                }
                body += "</tbody></table>";
            }
            if (noserverinfos.Count() > 0)
            {
                body += "\n<br /><div style'width:100%;'><span>Les serveurs ci-dessous ne sont pas en état critique de réplication; " +
                        "Cependant, ils ont présentés quelques défauts minimes:</span><br /><br />";

                body += "<table style='position:relative;width:100%;'><thead>" +
                "<tr style='position:relative;width:100%;height:35px;background-color:#0197bc'><th>Serveur</th><th>Durée d'absence</th><th>Ping</th>" +
                "<th>Etat</th><th>Origine</th><th>Site</th><th>IdSite</th><th>Erreurs</th><th>Détails</th></tr></thead><tbody>";
                for (int index = 0; index < noserverinfos.Count(); index++)
                {
                    string codeerror = noserverinfos[index].Errors.Split('(')[1];
                    codeerror = codeerror.Split(')')[0];
                    body += "<tr style='position:relative;width:100%;min-height:35px;border:1px solid #000;text-align:center;'>" +
                        "<td style='position:relative;padding-left:1px;width:8%;border:1px solid #000;'>" + noserverinfos[index].Server_Result.AD_Server.Name + "</td>" +
                        "<td style='position:relative;padding-left:1px;width:11%;border:1px solid #000;'>" + noserverinfos[index].Absence + "</td>" +
                        "<td style='position:relative;padding-left:1px;width:5%;border:1px solid #000;'>" + noserverinfos[index].Server_Result.Ping + "</td>" +
                        "<td style='position:relative;padding-left:1px;width:6%;border:1px solid #000;'>" + noserverinfos[index].Server_Result.AD_Server.Status + "</td>" +
                        "<td style='position:relative;padding-left:1px;width:5%;border:1px solid #000;'>" + noserverinfos[index].Origin + "</td>" +
                        "<td style='position:relative;padding-left:1px;width:15%;border:1px solid #000;'>" + noserverinfos[index].Server_Result.AD_Server.Site + "</td>" +
                        "<td style='position:relative;padding-left:1px;width:20%;border:1px solid #000;'>" + noserverinfos[index].Server_Result.AD_Server.IdSite + "</td>" +
                        "<td style='position:relative;padding-left:1px;width:5%;border:1px solid #000;'>" + codeerror + "</td>" +
                        "<td style='position:relative;padding-left:1px;width:25%;border:1px solid #000;'>" + noserverinfos[index].Errors.Substring(codeerror.Length + 2) + "</td></tr>";
                }
                body += "</tbody></table>";
            }
            body += "<br />" + Options + "</div>";

            email.Subject = "Résultat état des réplications pour la journée du : ";
            DateTime Today = DateTime.Today;
            email.Subject += DateTime.Now.ToString("dd") + "/" + DateTime.Now.ToString("MM") + "/" + DateTime.Now.ToString("yyyy"); ;
            email.Body = body;

            report.TotalErrors = totalErrors;
            report.FatalErrors = fatalErrors;
            TimeSpan Duration = DateTime.Now.Subtract(report.DateTime);
            report.Duration = Duration;
            if (ModelState.IsValid)
            {
                db.Entry(report).State = System.Data.Entity.EntityState.Modified;
                db.Entry(email).State = System.Data.Entity.EntityState.Modified;
                db.SaveChanges();
            }
            results.Add("Report", report.Id.ToString());
            results.Add("Email", email.Id.ToString());
            results.Add("Errors", errors);
            return Json(results, JsonRequestBehavior.AllowGet);
        }

        public string GetMaxDuration(string duration1, string duration2)
        {
            //Format(12d.05h:02m:20s)
            int days1 = 0, hours1 = 0, minutes1 = 0, seconds1 = 0;
            int days2 = 0, hours2 = 0, minutes2 = 0, seconds2 = 0;

            //PARSING FIRST DURATION
            if (duration1.IndexOf("d") != -1)
            {
                Int32.TryParse(duration1.Substring(0, 2).Trim(), out days1);
            }
            if (duration1.IndexOf("h") != -1)
            {
                Int32.TryParse(duration1.Substring(duration1.IndexOf("h") - 2, 2).Trim(), out hours1);
            }
            if (duration1.IndexOf("m") != -1)
            {
                Int32.TryParse(duration1.Substring(duration1.IndexOf("m") - 2, 2).Trim(), out minutes1);
            }
            if (duration1.IndexOf("s") != -1)
            {
                if (duration1.Length == 2)
                {
                    Int32.TryParse(duration1.Substring(0, 1).Trim(), out seconds1);
                }
                else
                {
                    Int32.TryParse(duration1.Substring(duration1.Length - 3, 2).Trim(), out seconds1);
                }
            }
            TimeSpan time1 = new TimeSpan(days1, hours1, minutes1, seconds1);

            //PARSING SECOND DURATION
            if (duration2.IndexOf("d") != -1)
            {
                Int32.TryParse(duration2.Substring(0, 2).Trim(), out days2);
            }
            if (duration2.IndexOf("h") != -1)
            {
                Int32.TryParse(duration2.Substring(duration2.IndexOf("h") - 2, 2).Trim(), out hours2);
            }
            if (duration2.IndexOf("m") != -1)
            {
                Int32.TryParse(duration2.Substring(duration2.IndexOf("m") - 2, 2).Trim(), out minutes2);
            }
            if (duration2.IndexOf("s") != -1)
            {
                if (duration2.Length == 2)
                {
                    Int32.TryParse(duration2.Substring(0, 1).Trim(), out seconds2);
                }
                else
                {
                    Int32.TryParse(duration2.Substring(duration2.Length - 3, 2).Trim(), out seconds2);
                }
            }
            TimeSpan time2 = new TimeSpan(days2, hours2, minutes2, seconds2);

            //COMPARE
            if (time1.CompareTo(time2) >= 0)
            {
                return duration1;
            }
            else
            {
                return duration2;
            }
        }

        public JsonResult CheckActiveDirectory()
        {
            AdReport report = db.AdReports.Create();
            report.DateTime = DateTime.Now;
            report.FatalErrors = 0;
            report.TotalErrors = 0;
            report.Module = HomeController.AD_MODULE;

            report.Author = User.Identity.Name;
            report.ResultPath = "";
            Email email = db.Emails.Create();
            email.Module = HomeController.AD_MODULE;
            email = Emails_Controller.SetRecipients(email, HomeController.AD_MODULE);

            report.Email = email;
            email.Report = report;
            DateTime Today = DateTime.Today;
            string SourceName = DateTime.Now.ToString("dd") + DateTime.Now.ToString("MM") + DateTime.Now.ToString("yy")
                + "-ReplSumDeltaBC1-";
            string FileName = SourceName + report.Id + ".txt";
            Dictionary<string, string[]> result = new Dictionary<string, string[]>();
            result.Add("response", null);
            result.Add("content", null);
            string[] response = new string[1];

            string PathDirectory = HomeController.AD_RESULTS_FOLDER;
            Process process = new Process();
            process.StartInfo.FileName = HomeController.BATCHES_FOLDER + "Repadmin_Launcher_v4.bat";
            process.StartInfo.UseShellExecute = false;
            process.StartInfo.RedirectStandardOutput = false;
            process.StartInfo.WorkingDirectory = PathDirectory;
            process.StartInfo.Arguments = report.Id.ToString();
            process.Start();
            process.WaitForExit();

            try
            {
                string[] lines = System.IO.File.ReadAllLines(PathDirectory + FileName, Encoding.Default);
                if (ModelState.IsValid)
                {
                    db.AdReports.Add(report);
                    report.ResultPath = PathDirectory + FileName;
                    TimeSpan Duration = (TimeSpan)DateTime.Now.Subtract(report.DateTime);
                    report.Duration = Duration;
                    db.SaveChanges();
                    System.IO.File.Copy(PathDirectory + FileName, PathDirectory + SourceName + report.Id + ".txt");
                    report.ResultPath = PathDirectory + SourceName + report.Id + ".txt";
                    db.Entry(report).State = System.Data.Entity.EntityState.Modified;
                    db.SaveChanges();
                    string[] ReportId = new string[1];
                    ReportId[0] = report.Id.ToString();
                    response[0] = "success";
                    result["response"] = response;
                    result["content"] = lines;
                    result.Add("Report", ReportId);
                    result.Add("Email", ReportId);

                    int reportNumber = db.AdReports.Count();
                    if (reportNumber > HomeController.AD_MAX_REPORT_NUMBER)
                    {
                        int reportNumberToDelete = reportNumber - HomeController.AD_MAX_REPORT_NUMBER;
                        AdReport[] reportsToDelete =
                            db.AdReports.OrderBy(id => id.Id).Take(reportNumberToDelete).ToArray();
                        foreach (AdReport toDeleteReport in reportsToDelete)
                        {
                            DeleteAdReport(toDeleteReport.Id);
                        }
                    }
                }
                Specific_Logging(new Exception("...."), "CheckActiveDirectory", 3);
                return Filter(Json(result, JsonRequestBehavior.AllowGet), report.Id);
            }
            catch (Exception exception)
            {
                Specific_Logging(exception, "CheckActiveDirectory");
                response[0] = "failed";
                string[] content = new string[1];
                content[0] = exception.Message;
                result["response"] = response;
                result["content"] = content;
                return Json(result, JsonRequestBehavior.AllowGet);
            }
        }

        public ActionResult GetReportStatistics()
        {
            AdReport[] reports = db.AdReports.ToArray();
            String[] reportsDate = new String[reports.Length];
            Object[] totalerrors = new Object[reports.Length];
            Object[] fatalerrors = new Object[reports.Length];
            int index = 0;
            foreach (AdReport report in reports)
            {
                reportsDate[index] = report.DateTime.ToString();
                totalerrors[index] = report.TotalErrors;
                fatalerrors[index] = report.FatalErrors;
                index++;
            }

            System.Drawing.Color[] ColumnColors = { };

            Highcharts chart = new Highcharts("chart")
                .InitChart(new Chart
                {
                    DefaultSeriesType = ChartTypes.Column,
                    MarginRight = 130,
                    MarginBottom = 25,
                    BorderWidth = 0,
                    BorderRadius = 15,
                    PlotBackgroundColor = null,
                    PlotShadow = false,
                    PlotBorderWidth = 0,
                    BackgroundColor = new BackColorOrGradient(ColorTranslator.FromHtml("#ded9d4")),
                    ClassName = "Report"
                })
                .SetOptions(new GlobalOptions
                {
                    Colors = new[]
                                         {
                                             ColorTranslator.FromHtml("#DDDF0D"),
                                             ColorTranslator.FromHtml("#7798BF"),
                                             ColorTranslator.FromHtml("#55BF3B"),
                                             ColorTranslator.FromHtml("#DF5353"),
                                             ColorTranslator.FromHtml("#DDDF0D"),
                                             ColorTranslator.FromHtml("#aaeeee"),
                                             ColorTranslator.FromHtml("#ff0066"),
                                             ColorTranslator.FromHtml("#eeaaee")

                                         }
                })
                .SetTitle(new Title
                {
                    Text = "Contrôleurs défaillant",
                    X = -20
                })
                .SetSubtitle(new Subtitle
                {
                    Text = "Le nombre de contrôleurs défaillant ayant été détectés ces derniers jours",
                    X = -20
                })
                .SetXAxis(new XAxis
                {
                    Categories = reportsDate
                })
                .SetYAxis(new YAxis
                {
                    Title = new YAxisTitle { Text = "Nombre de contrôleurs" },
                    PlotLines = new[]
                            {
                                new YAxisPlotLines
                                    {
                                        Value = 0,
                                        Width = 1,
                                        Color = ColorTranslator.FromHtml("#f00")
                                    }
                            }
                })
                .SetTooltip(new Tooltip
                {
                    Formatter = @"function() {
                                        return '<b>'+ this.series.name +'</b><br/>'+
                                    this.x +': '+ this.y +' contrôleurs défaillant';
                                }",
                    BorderWidth = 0,
                    Shadow = false,
                    BackgroundColor = new BackColorOrGradient(ColorTranslator.FromHtml("#fff")),
                    BorderColor = ColorTranslator.FromHtml("transparent"),
                    //Enabled = false

                })
                .SetLegend(new Legend
                {
                    Layout = Layouts.Vertical,
                    Align = HorizontalAligns.Right,
                    VerticalAlign = VerticalAligns.Top,
                    X = -10,
                    Y = 100,
                    BorderWidth = 0
                })
                .SetSeries(new[]
                    {
                        new Series { Name = "Serveurs défaillant", Data = new Data(totalerrors),Color = ColorTranslator.FromHtml("#e75114") },
                        new Series { Name = "Serveurs en défaillance critique", Data = new Data(fatalerrors),Color = ColorTranslator.FromHtml("#ff3f3f") },
                    }
                );

            return View(chart);

        }

        public string Purge()
        {
            string message = "";
            List<AdReport> reports = db.AdReports.Where(rep => rep.Duration == null || rep.ResultPath == null).ToList();
            foreach (AdReport report in reports)
            {
                message += "Rapport " + report.DateTime + " supprimé";
                Email email = (report.Email != null) ? report.Email : null;
                List<FaultyServer_Report> faultyserverreports = report.FaultyServer_Reports.ToList();
                foreach (FaultyServer_Report faultyserverreport in faultyserverreports)
                {
                    db.FaultyServerReports.Remove(faultyserverreport);
                }
                db.SaveChanges();
                Reports_Controller.Delete(report.Id);
            }
            Specific_Logging(new Exception("...."), "Purge", 3);
            return message;
        }

        public string DeleteFaultyServer(int id)
        {
            try
            {
                FaultyServer server = db.FaultyServers.Find(id);
                FaultyServer_Report[] faultyserverreports = server.FaultyServer_Reports.ToArray();

                foreach (FaultyServer_Report faultyserverreport in faultyserverreports)
                {
                    db.FaultyServerReports.Remove(faultyserverreport);
                    if (faultyserverreport.AdReport != null)
                    {
                        DeleteAdReport(faultyserverreport.AdReport.Id);
                    }
                }
                db.FaultyServers.Remove(server);
                db.SaveChanges();
                Specific_Logging(new Exception("...."), "DeleteFaultyServer", 3);
                return "Toutes les références sur ce contrôleur ont été supprimées.";
            }
            catch (Exception exception)
            {
                Specific_Logging(exception, "DeleteFaultyServer");
                return "Une erreur est survenue lors de la suppression:\n" + exception.Message;
            }
        }

        public string LaunchTestError()
        {
            AdReport report = db.AdReports.OrderByDescending(id => id.Id).First();
            Email email = report.Email;

            //TEST USER AND PASSWORD
            string domain = HomeController.DEFAULT_DOMAIN_IMPERSONNATION;
            string username = HomeController.DEFAULT_USERNAME_IMPERSONNATION;
            string password = McoUtilities.Decrypt(HomeController.DEFAULT_PASSWORD_IMPERSONNATION);
            try
            {
                password = Request.Form["password"].ToString();
                domain = Request.Form["domain"].ToString();
                username = Request.Form["username"].ToString();
            }
            catch { }

            IntPtr userToken = IntPtr.Zero;
            bool success = McoUtilities.LogonUser(
                    username,
                    domain,
                    password,
                    (int)McoUtilities.LogonType.LOGON32_LOGON_INTERACTIVE, //2
                    (int)McoUtilities.LogonProvider.LOGON32_PROVIDER_DEFAULT, //0
                    out userToken);

            if (!McoUtilities.IsValidLoginPassword(domain + @"\" + username, password))
            {
                return "Mauvaise combinaison utilisateur/mot de passe!";
            }
            //PARSING OF THE REPORT
            if (report == null || email == null)
            {
                return HttpNotFound().ToString();
            }
            string[] tables = email.Body.Split(new string[] { "<table" }, StringSplitOptions.RemoveEmptyEntries);
            string result = "";

            foreach (string table in tables)
            {
                try
                {
                    string[] headers = table.Split(new string[] { "<tbody>" }, StringSplitOptions.RemoveEmptyEntries);
                    if (headers.Length >= 1)
                    {
                        string[] lines = headers[1].Split(new string[] { "<tr" }, StringSplitOptions.RemoveEmptyEntries);
                        foreach (string line in lines)
                        {
                            string[] columns = line.Split(new string[] { "<td" }, StringSplitOptions.RemoveEmptyEntries);
                            string error = columns[8].Split('>')[1];
                            error = error.Split('<')[0];
                            string server = columns[1].Split('>')[1];
                            server = server.Split('<')[0];
                            if (error.Trim() == "8606")
                            {
                                result += "------------------------- \n";
                                result += TestServerError(error, server, username, domain, password) + "\n";
                            }
                        }
                    }
                }
                catch { }
            }
            if (result == "")
            {
                return "L'erreur 8606 n'a été détectée sur aucun serveur du rapport du " + email.Report.DateTime.ToString();
            }
            return result;
        }

        public string TestServerError(string error, string servername, string username, string domain, string password)
        {
            string result = "";
            if (error.Trim() == "8606")
            {
                string partialserverFolder = @"\\" + servername + @"\d$\temp";
                string serverFolder = @"\\" + servername + @"\d$\temp\";
                string localFolder = HomeController.BATCHES_FOLDER;
                IntPtr userToken = IntPtr.Zero;
                bool success = McoUtilities.LogonUser(
                    username,
                    domain,
                    password,
                    (int)McoUtilities.LogonType.LOGON32_LOGON_INTERACTIVE, //2
                    (int)McoUtilities.LogonProvider.LOGON32_PROVIDER_DEFAULT, //0
                    out userToken);

                if (!success)
                {
                    return "Mauvaise combinaison utilisateur/mot de passe!";
                }

                using (WindowsIdentity.Impersonate(userToken))
                {
                    //CONNECTION FOR REMOTE 
                    System.Security.Principal.WindowsImpersonationContext impersonationContext;
                    impersonationContext = ((System.Security.Principal.WindowsIdentity)User.Identity).Impersonate();
                    try
                    {
                        Process process = new Process();
                        process.StartInfo.FileName = "cmd.exe";
                        process.StartInfo.LoadUserProfile = true;
                        process.StartInfo.UseShellExecute = false;
                        process.StartInfo.RedirectStandardOutput = false;
                        process.StartInfo.RedirectStandardInput = true;

                        using (UNC_ACCESSOR)
                        {
                            if (UNC_ACCESSOR.NetUseWithCredentials(partialserverFolder, username, domain, password))
                            {
                                //The values of username, domain and password are the same as the logged user
                                //The rights used here are those of the logged users
                            }
                            else
                            {
                                UNC_ACCESSOR.NetUseWithCredentials(partialserverFolder,
                                    HomeController.DEFAULT_USERNAME_IMPERSONNATION,
                                    HomeController.DEFAULT_DOMAIN_IMPERSONNATION,
                                    McoUtilities.Decrypt(HomeController.DEFAULT_PASSWORD_IMPERSONNATION));
                                success = McoUtilities.LogonUser(
                                        HomeController.DEFAULT_USERNAME_IMPERSONNATION,
                                        HomeController.DEFAULT_DOMAIN_IMPERSONNATION,
                                        McoUtilities.Decrypt(HomeController.DEFAULT_PASSWORD_IMPERSONNATION),
                                        (int)McoUtilities.LogonType.LOGON32_LOGON_INTERACTIVE, //2
                                        (int)McoUtilities.LogonProvider.LOGON32_PROVIDER_DEFAULT, //0
                                        out userToken);

                                if (!success)
                                {
                                    return "Impossible de récupérer les droits d'accès de l'utilisateur";
                                }
                                //Use of BCNDOMAIN\srvexploit's account rights and permissions
                            }
                            System.IO.File.Copy(localFolder + "GetEventLog.txt",
                                    serverFolder + "GetEventLog_" + servername + ".ps1", true);

                            process.StartInfo.UserName = username;
                            process.StartInfo.Domain = domain;
                            var secure = new System.Security.SecureString();
                            foreach (char c in password)
                            {
                                secure.AppendChar(c);
                            }
                            process.StartInfo.Password = secure;

                            process.Start();
                            process.StandardInput.WriteLine("pushd " + serverFolder);
                            process.StandardInput.WriteLine("powershell GetEventLog_" + servername + ".ps1");
                            process.StandardInput.WriteLine("exit");
                            process.WaitForExit();

                            string resultFilename = HomeController.AD_RESULTS_FOLDER + "EventLogOK_" + servername + ".csv";
                            string foundedSources = "";

                            //FINAL FILE PARSING
                            try
                            {
                                System.IO.File.Copy(serverFolder + "EventLogOK.csv",
                                resultFilename, true);

                                string[] lines = System.IO.File.ReadAllLines(resultFilename, Encoding.Default);
                                string newFileContent = "";
                                foreach (string line in lines)
                                {
                                    if ((line.IndexOf("Server") != -1) && (line.IndexOf("Source") != -1)
                                        && (line.IndexOf("Object") != -1))
                                    {
                                        continue;
                                    }
                                    if (line.Trim() != "")
                                    {
                                        string[] infos = line.Split(';');
                                        if (infos.Length >= 4)
                                        {
                                            if (foundedSources.IndexOf(infos[2]) != -1)
                                            {
                                                continue;
                                            }
                                            else
                                            {
                                                foundedSources += infos[2] + " ";
                                            }

                                            newFileContent += "repadmin /removelingeringobjects ";
                                            infos[0] = infos[0].Substring(1, infos[0].Length - 2);
                                            newFileContent += infos[0] + " ";
                                            int source = infos[2].IndexOf("._");
                                            if (source != -1)
                                            {
                                                infos[2] = infos[2].Substring(0, source);
                                                infos[2] = infos[2].Substring(1);
                                                newFileContent += infos[2] + " ";
                                            }
                                            int forest = infos[3].IndexOf("DC=ForestDnsZones");
                                            if (forest != -1)
                                            {
                                                infos[3] = infos[3].Substring(forest, infos[3].Length - forest);
                                                infos[3] = infos[3].Substring(0, infos[3].Length - 1);
                                                newFileContent += infos[3];
                                            }
                                            newFileContent += "\r\n";
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
                                }
                                string finalFilename = HomeController.AD_RESULTS_FOLDER + "CleanLingering." + servername + ".cmd";
                                System.IO.File.WriteAllText(finalFilename, newFileContent);

                                System.IO.File.Copy(finalFilename,
                                    serverFolder + "\\CleanLingering." + servername + ".cmd", true);
                            }
                            catch (Exception exception)
                            {
                                result = exception.Message;
                                goto ERROR_ON_LAUNCHING;
                            }
                        }
                    }
                    catch (Exception exception)
                    {
                        result = exception.Message;
                        goto ERROR_ON_LAUNCHING;
                    }

                    impersonationContext.Undo();
                }
                Specific_Logging(new Exception("...."), "TestServerError " + servername, 2);
                return "L'opération a été exécutée sur " + servername;
            }
            else
            {
                return "\n";
            }
        ERROR_ON_LAUNCHING:
            {
                result = "L'opération n'a pas été correctement exécutée sur " + servername + ": \n" + result;
                return result;
            }
        }

        public class VirtualizedAdServer
        {
            public ServersController.VirtualizedServer_Result Server_Result { get; set; }
            public string Absence { get; set; }
            public string Errors { get; set; }
            public int Fail { get; set; }
            public string Origin { get; set; }

            public VirtualizedAdServer()
            {
                this.Server_Result = new ServersController.VirtualizedServer_Result();
                this.Absence = this.Errors = this.Origin = "";
                this.Fail = 0;
            }

            public VirtualizedAdServer(FaultyServer server)
            {
                this.Server_Result = new ServersController.VirtualizedServer_Result(server);
                this.Server_Result.AD_Server = server;
                this.Absence = this.Errors = this.Origin = "";
                this.Fail = 0;
            }

            public static VirtualizedAdServer GetServerInformations(Dictionary<int, ServersController.VirtualizedServer> FOREST, ReftechServers[] REFTECH_SERVERS, FaultyServer server, bool ping = true)
            {
                VirtualizedAdServer virtual_server = new VirtualizedAdServer(server);
                virtual_server.Server_Result = ServersController.GetServerInformations(FOREST, REFTECH_SERVERS, server, HomeController.AD_MODULE, ping);
                return virtual_server;
            }
        }

        public JsonResult ExecuteSchedule(int id)
        {
            AdSchedule schedule = db.AdSchedules.Find(id);
            if (schedule == null)
            {
                Specific_Logging(new Exception("...."), "ExecuteSchedule");
                return null;
            }
            schedule.State = "En cours";
            if (ModelState.IsValid)
            {
                db.Entry(schedule).State = System.Data.Entity.EntityState.Modified;
                db.SaveChanges();
            }
            AdReport report = db.AdReports.Create();
            report.DateTime = DateTime.Now;
            report.FatalErrors = 0;
            report.TotalErrors = 0;
            report.Module = HomeController.AD_MODULE;
            report.ScheduleId = schedule.Id;
            report.Author = HomeController.SYSTEM_IDENTITY;
            report.ResultPath = "";
            Email email = db.Emails.Create();
            email.Module = HomeController.AD_MODULE;

            email = Emails_Controller.SetRecipients(email, HomeController.AD_MODULE);

            report.Email = email;
            email.Report = report;
            DateTime Today = DateTime.Today;
            string SourceName = DateTime.Now.ToString("dd") + DateTime.Now.ToString("MM") + DateTime.Now.ToString("yy")
                + "-ReplSumDeltaBC1-";
            string FileName = SourceName + report.Id + ".txt";
            Dictionary<string, string[]> result = new Dictionary<string, string[]>();
            result.Add("response", null);
            result.Add("content", null);
            string[] response = new string[1];

            string PathDirectory = HomeController.AD_RESULTS_FOLDER;
            Process process = new Process();
            process.StartInfo.FileName = HomeController.BATCHES_FOLDER + "RepadminLauncher.bat";
            process.StartInfo.UseShellExecute = false;
            process.StartInfo.RedirectStandardOutput = false;
            process.StartInfo.WorkingDirectory = PathDirectory;
            process.StartInfo.Arguments = report.Id.ToString();
            process.Start();
            process.WaitForExit();

            try
            {
                string[] lines = System.IO.File.ReadAllLines(PathDirectory + FileName, Encoding.Default);
                if (ModelState.IsValid)
                {
                    db.AdReports.Add(report);
                    report.ResultPath = PathDirectory + FileName;
                    TimeSpan Duration = (TimeSpan)DateTime.Now.Subtract(report.DateTime);
                    report.Duration = Duration;
                    db.SaveChanges();
                    System.IO.File.Copy(PathDirectory + FileName, PathDirectory + SourceName + report.Id + ".txt");
                    report.ResultPath = PathDirectory + SourceName + report.Id + ".txt";
                    db.Entry(report).State = System.Data.Entity.EntityState.Modified;
                    db.SaveChanges();
                    string[] ReportId = new string[1];
                    ReportId[0] = report.Id.ToString();
                    response[0] = "success";
                    result["response"] = response;
                    result["content"] = lines;
                    result.Add("Report", ReportId);
                    result.Add("Email", ReportId);

                    int reportNumber = db.AdReports.Count();
                    if (reportNumber > HomeController.AD_MAX_REPORT_NUMBER)
                    {
                        int reportNumberToDelete = reportNumber - HomeController.AD_MAX_REPORT_NUMBER;
                        AdReport[] reportsToDelete =
                            db.AdReports.OrderBy(ide => ide.Id).Take(reportNumberToDelete).ToArray();
                        foreach (AdReport toDeleteReport in reportsToDelete)
                        {
                            DeleteAdReport(toDeleteReport.Id);
                        }
                    }
                }
                Specific_Logging(new Exception("...."), "CheckActiveDirectory", 3);
                JsonResult filter = Filter(Json(result, JsonRequestBehavior.AllowGet), report.Id);
                Emails_Controller.AutoSend(email.Id);
                schedule.State = (schedule.Multiplicity != "Une fois") ? "Planifié" : "Terminé";
                schedule.NextExecution = Schedules_Controller.GetNextExecution(schedule);
                schedule.Executed++;
                if (ModelState.IsValid)
                {
                    db.Entry(schedule).State = System.Data.Entity.EntityState.Modified;
                    db.SaveChanges();
                }
                return filter;
            }
            catch (Exception exception)
            {
                Specific_Logging(exception, "CheckActiveDirectory");
                response[0] = "failed";
                string[] content = new string[1];
                content[0] = exception.Message;
                result["response"] = response;
                result["content"] = content;
                schedule.State = (schedule.Multiplicity != "Une fois") ? "Planifié" : "Terminé";
                schedule.Executed++;
                if (ModelState.IsValid)
                {
                    db.Entry(schedule).State = System.Data.Entity.EntityState.Modified;
                    db.SaveChanges();
                }
                return Json(result, JsonRequestBehavior.AllowGet);
            }
        }

        private void Specific_Logging(Exception exception, string action, int level = 0)
        {
            string author = "UNKNOWN";
            if (User != null && User.Identity != null && User.Identity.Name != " ")
            {
                author = User.Identity.Name;
            }
            McoUtilities.Specific_Logging(exception, action, HomeController.AD_MODULE, level, author);
        }

        public string HasUnachieviedReport()
        {
            int unachievied = db.AdReports.Where(rep => rep.Duration == null || rep.ResultPath == null).Count();
            if (unachievied > 0)
            {
                return "Il y a " + unachievied + " rapport(s) inachevé(s)\nVoulez-vous les supprimer?";
            }
            return "OK";
        }

    }
}
