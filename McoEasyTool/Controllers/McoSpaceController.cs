using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using Excel = Microsoft.Office.Interop.Excel;

namespace McoEasyTool.Controllers
{
    [AllowAnonymous]
    public class McoSpaceController : Controller
    {
        private DataModelContainer db = new DataModelContainer();
        private ReportsController Reports_Controller = new ReportsController();
        private EmailsController Emails_Controller = new EmailsController();
        private SchedulesController Schedules_Controller = new SchedulesController();
        private ServersController Servers_Controller = new ServersController();
        private AccountsController Accounts_Controller = new AccountsController();
        private McoUtilities.UNCAccessWithCredentials UNC_ACCESSOR = new McoUtilities.UNCAccessWithCredentials();
        private static Excel.Workbook MyWorkbook = null;
        private static Excel.Application MyApplication = null;
        private static Excel.Worksheet MySheet = null;

        public ActionResult Home()
        {
            ViewBag.SPACE_DESC_0 = McoUtilities.GetModuleDescription(HomeController.SPACE_MODULE, 0);
            ViewBag.SPACE_DESC_1 = McoUtilities.GetModuleDescription(HomeController.SPACE_MODULE, 1);
            ViewBag.SPACE_DESC_2 = McoUtilities.GetModuleDescription(HomeController.SPACE_MODULE, 2);
            ViewBag.SPACE_DESC_3 = McoUtilities.GetModuleDescription(HomeController.SPACE_MODULE, 3);
            ViewBag.SPACE_DESC_4 = McoUtilities.GetModuleDescription(HomeController.SPACE_MODULE, 4);
            return View();
        }

        public ActionResult DisplaySchedules()
        {
            return View(db.SpaceSchedules.OrderBy(name => name.TaskName).ToList());
        }

        public ActionResult DisplayReports()
        {
            return View(db.SpaceReports.OrderByDescending(report => report.Id).ToList());
        }

        public ActionResult DisplayScheduleReports(int id)
        {
            SpaceSchedule schedule = db.SpaceSchedules.Find(id);
            object[] boundaries =
                McoUtilities.GetIdValues<SpaceSchedule>(schedule, HomeController.OBJECT_ATTR_ID);
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
            ICollection<SpaceReport> reports = db.SpaceReports.Where(report => report.ScheduleId == id).ToList();
            return View(reports.OrderByDescending(report => report.Id).ToList());
        }

        public ActionResult DisplayFailedServers()
        {
            SpaceReport report = db.SpaceReports.LastOrDefault();

            List<SpaceServer_Report> spaceserverreports = null;
            try
            {
                spaceserverreports = db.SpaceServerReports
                    .Where(state => state.State != "OK").Where(spa => spa.SpaceReportId == report.Id)
                    .Distinct().ToList();
            }
            catch { }

            if (spaceserverreports == null)
            {
                return View();
            }
            string list = "";
            foreach (SpaceServer_Report serverreport in spaceserverreports)
            {
                list += serverreport.SpaceServer.Id + ",";
            }

            if (spaceserverreports == null)
            {
                return View();
            }

            return View(spaceserverreports.OrderByDescending(rep => rep.SpaceReportId).ToList());
        }

        public ActionResult DisplayRecipients()
        {
            return View(db.Recipients.Where(rec => rec.Module == HomeController.SPACE_MODULE).ToList());
        }

        public ActionResult DisplayImporter()
        {
            return View();
        }

        public ActionResult DisplaySpaceServers()
        {
            return View(db.SpaceServers.OrderBy(spa => spa.Name).ToList());
        }

        public ActionResult DisplayChecker()
        {
            return View(db.SpaceServers.OrderBy(spa => spa.Name).ToList());
        }

        public ActionResult UploadInitFile()
        {
            HttpPostedFileBase file = Request.Files[0];
            if (file != null)
                file.SaveAs(HomeController.SPACE_INIT_FILE);
            ViewBag.Message = Import(true);
            return DisplaySpaceServers();
        }

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
                    string taskname = HomeController.SPACE_MODULE + " AutoCheck ";

                    SpaceSchedule schedule = db.SpaceSchedules.Create();
                    schedule.CreationTime = DateTime.Now;
                    schedule.NextExecution = scheduled;
                    schedule.Generator = User.Identity.Name;
                    schedule.Multiplicity = multiplicity;
                    schedule.Executed = 0;
                    schedule.State = "Planifié";
                    schedule.Module = HomeController.SPACE_MODULE;
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
                        Specific_Logging(new Exception("...."), "CreateSchedule", 3);
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
            Specific_Logging(new Exception("...."), "CreateSchedule", 2);
            return result;
        }

        public string EditSchedule(int id)
        {
            SpaceSchedule schedule = db.SpaceSchedules.Find(id);
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
                    string taskname = HomeController.SPACE_MODULE + " AutoCheck ";

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
                        db.Entry(schedule).State = System.Data.Entity.EntityState.Modified;
                        db.SaveChanges();
                        Specific_Logging(new Exception("...."), "EditSchedule", 3);
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
            Specific_Logging(new Exception("...."), "EditSchedule", 2);
            return result;
        }

        public string DeleteSchedule(int id)
        {
            SpaceSchedule schedule = db.SpaceSchedules.Find(id);

            if (schedule.Reports.Count != 0)
            {
                List<SpaceReport> reports = db.SpaceReports.Where(rep => rep.ScheduleId == schedule.Id).ToList();
                foreach (Report report in reports)
                {
                    DeleteSpaceReport(report.Id);
                    db.Entry(schedule).State = System.Data.Entity.EntityState.Modified;
                    db.SaveChanges();
                }
            }

            string result = Schedules_Controller.Delete(schedule);
            schedule = db.SpaceSchedules.Find(id);
            if (result == "La tâche a été correctement supprimée")
            {
                db.SpaceSchedules.Remove(schedule);
                db.SaveChanges();
            }
            Specific_Logging(new Exception("...."), "DeleteSchedule", 3);
            return result;
        }

        public class VirtualSpaceServerResult
        {
            public string Name { get; set; }
            public string Threshold { get; set; }
            public string State { get; set; }
            public string CellColor { get; set; }
            public int ReportId { get; set; }
            public List<VirtualPartitionResult> Disks { get; set; }

            public VirtualSpaceServerResult(SpaceServer_Report serverreport)
            {
                this.Name = serverreport.SpaceServer.Name.Trim();
                this.State = serverreport.State.Trim();
                this.ReportId = serverreport.SpaceReportId;
                this.CellColor = serverreport.SpaceServer.CellColor.Trim();
                Disks = VirtualPartitionResult.GetPartitions(serverreport.SpaceServer);
            }

            public VirtualPartitionResult GetPartition(string name)
            {
                foreach (VirtualPartitionResult partition in this.Disks)
                {
                    if (partition.Name.ToUpper().Trim() == name.ToUpper().Trim())
                    {
                        return partition;
                    }
                }
                return null;
            }

            public void SetPartitionInfos(string name, string free, string status)
            {
                try
                {
                    VirtualPartitionResult partition = GetPartition(name);
                    int index = this.Disks.IndexOf(partition);
                    partition.Free = free.Trim();
                    partition.State = status.Trim();
                    string color = (status != "OK") ? " color:#ff2f00;font-weight:bold;" : "";
                    partition.Display = "<td title='" + this.Name +
                        "_" + partition.Name + "' style='background-color:" +
                        this.CellColor + ";" + color + "'>" + partition.Free + "</td>";
                    this.Disks[index] = partition;
                }
                catch { }
            }
        }

        public class VirtualPartitionResult
        {
            public string Name { get; set; }
            public string Threshold { get; set; }
            public string Free { get; set; }
            public string State { get; set; }
            public string Display { get; set; }

            public VirtualPartitionResult(string name, string threshold)
            {
                this.Name = name.Trim();
                this.Threshold = threshold.Trim();
            }

            public static List<VirtualPartitionResult> GetPartitions(SpaceServer server)
            {
                List<VirtualPartitionResult> partitions = new List<VirtualPartitionResult>();
                string[] parser = server.Disks.Split(new string[] { "; " }, StringSplitOptions.RemoveEmptyEntries);
                for (int index = 0; index < parser.Length; index++)
                {
                    string[] infos = parser[index].Split(new string[] { "|" }, StringSplitOptions.RemoveEmptyEntries);
                    string disk = infos[0].Substring(infos[0].IndexOf("=") + 1).Trim();
                    string threshold = infos[1].Substring(infos[1].IndexOf("=") + 1).Trim();
                    threshold = threshold.Split(' ')[0].Trim();
                    string partition = (server.IsShare) ? disk.Substring(disk.LastIndexOf(@"\")) : "Disque " + disk.Trim();
                    partitions.Add(new VirtualPartitionResult(partition, threshold));
                }
                return partitions;
            }

            public override string ToString()
            {
                if (this.Display != null)
                {
                    return this.Display;
                }
                else
                {
                    return "<td>N/A</td>";
                }
            }
        }

        public string BuildEmail(int id)
        {
            Email email = db.Emails.Find(id);
            if (email == null)
            {
                return HttpNotFound().ToString();
            }
            SpaceReport the_report = (SpaceReport)email.Report;
            List<SpaceServer_Report> the_serverreports = the_report.SpaceServer_Reports.ToList();
            List<SpaceServer> servers = new List<SpaceServer>();
            List<VirtualSpaceServerResult> servers_results = new List<VirtualSpaceServerResult>();
            foreach (SpaceServer_Report serverreport in the_serverreports)
            {
                if (!servers.Contains(serverreport.SpaceServer))
                {
                    servers.Add(serverreport.SpaceServer);
                }
            }
            servers = servers.OrderBy(ser => ser.Name).ToList();
            string body = "<br/><style>th{border:1px solid #fff;font-weight:bold;position:relative;}td{border:1px solid #ccc;text-align:center;}tr:hover>td,tr:hover>td a{cursor:pointer;color:#fff;background-color:#68b3ff;}</style>";
            body += "<table style='position:relative;width:100%;' cellpadding='0' cellspacing='0'>" +
                "<thead><tr style='position:relative;text-align:center;width:100%;background-color:#dcdbdb;'>" +
                    "<th style='position:relative;text-align:center;font-weight:bold;border:1px solid #fff;'>Serveurs</th>";

            string second_head = "<tr><th>Partitions</th>";
            string third_head = "<tr style='background-color:#ffc000;'><th>Seuils</th>";
            foreach (SpaceServer server in servers)
            {
                string[] parser = server.Disks.Split(new string[] { "; " }, StringSplitOptions.RemoveEmptyEntries);
                for (int index = 0; index < parser.Length; index++)
                {
                    string[] infos = parser[index].Split(new string[] { "|" }, StringSplitOptions.RemoveEmptyEntries);
                    string disk = infos[0].Substring(infos[0].IndexOf("=") + 1).Trim();
                    string threshold = infos[1].Substring(infos[1].IndexOf("=") + 1).Trim();
                    threshold = threshold.Split(' ')[0].Trim();
                    string partition = (server.IsShare) ? disk.Substring(disk.LastIndexOf(@"\")) : "Disque " + disk;
                    second_head += "<th style='background-color:" + server.CellColor + ";'>" + partition + "</th>";
                    third_head += "<th>" + threshold + " To</th>";
                }
                body += "<th style='background-color:" + server.CellColor + ";' colspan='" + parser.Length + "'>" + server.Name + "</th>";
            }
            second_head += "</tr>";
            third_head += "</tr>";
            body += "</tr>" + second_head + third_head + "</thead><tbody>";

            List<SpaceReport> reports = db.SpaceReports.OrderByDescending(rep => rep.Id).Take(7).ToList();
            reports = reports.OrderBy(rep => rep.DateTime).ToList();
            foreach (SpaceReport report in reports)
            {
                List<SpaceServer_Report> serverreports = report.SpaceServer_Reports.OrderBy(ser => ser.SpaceServer.Name).ToList();
                int current_index = 0;
                foreach (SpaceServer_Report serverreport in serverreports)
                {
                    VirtualSpaceServerResult server_result = new VirtualSpaceServerResult(serverreport);
                    string[] parser = serverreport.Details.Split(new string[] { "; " }, StringSplitOptions.RemoveEmptyEntries);
                    for (int index = 0; index < parser.Length; index++)
                    {
                        string[] infos = parser[index].Split(new string[] { "|" }, StringSplitOptions.RemoveEmptyEntries);
                        string free = infos[1].Substring(infos[1].IndexOf("=") + 1).Trim();
                        string disk = infos[0].Substring(infos[0].IndexOf("=") + 1).Trim();
                        string partition = (serverreport.SpaceServer.IsShare) ? disk.Substring(disk.LastIndexOf(@"\")) : "Disque " + disk;
                        string status = infos[2].Trim();
                        server_result.SetPartitionInfos(partition, free, status);
                    }
                    servers_results.Add(server_result);
                }
            }

            foreach (SpaceReport report in reports)
            {
                List<VirtualSpaceServerResult> results =
                    servers_results.Where(ser => ser.ReportId == report.Id)
                    .OrderBy(ser => ser.Name).ToList();
                body += "<tr><td>" + report.DateTime.ToString() + "</td>";
                foreach (SpaceServer server in servers)
                {
                    if (results.Where(ser => ser.Name == server.Name).Count() == 1)
                    {
                        VirtualSpaceServerResult virtual_server =
                            results.Where(ser => ser.Name == server.Name).FirstOrDefault();
                        foreach (VirtualPartitionResult partition in virtual_server.Disks)
                        {
                            body += partition.ToString();
                        }
                    }
                    else
                    {
                        string[] parser = server.Disks.Split(new string[] { "; " }, StringSplitOptions.RemoveEmptyEntries);
                        for (int index = 0; index < parser.Length; index++)
                        {
                            body += "<td>N/A</td>";
                        }
                    }
                }
                body += "</tr>";
            }
            body += "</tbody></table>";
            body += "<br/><br/><span>Les entrées en rouge correspondent à celles dont l'espace libre est inférieur au seuil.</span><br/>";
            body += "<span>Ne sont référencées dans le tableau que les entrées correspondant aux sept (07) derniers rapports.</span>";

            email.Body = body;
            email.Subject = "Resultat check Capacity Planning " + email.Report.DateTime.ToString();
            if (ModelState.IsValid)
            {
                db.Entry(email).State = System.Data.Entity.EntityState.Modified;
                db.SaveChanges();
            }
            return "BuildOK";
        }

        public string ReSendLastEmail(int id)
        {
            SpaceSchedule schedule = db.SpaceSchedules.Find(id);
            if (schedule != null)
            {
                if (db.Reports.Where(report => report.ScheduleId == id).Count() != 0)
                {
                    SpaceReport report = db.SpaceReports.Where(rep => rep.ScheduleId == id).OrderByDescending(rep => rep.Id).First();
                    return Reports_Controller.ReSend(report.Id);
                }
                Specific_Logging(new Exception("...."), "ReSendLastEmail", 3);
                return "Cette tâche planifiée n'a pour l'instant généré aucun rapport, ou alors ils ont été supprimés.";

            }
            Specific_Logging(new Exception("...."), "ReSendLastEmail", 2);
            return "Cette tâche planifiée n'a pas été trouvée dans la base de données.";
        }

        public string ViewEmails(int id)
        {
            SpaceReport report = db.SpaceReports.Find(id);
            if (report != null)
            {
                return Reports_Controller.ViewEmail(report.Id);
            }
            return "Ce rapport n'a pas été retrouvé dans la base de données.";
        }

        public string ReSendEmail(int id)
        {
            SpaceReport report = db.SpaceReports.Find(id);
            if (report != null)
            {
                return Reports_Controller.ReSend(report.Id);
            }
            Specific_Logging(new Exception("...."), "ReSendEmail", 2);
            return "Le rapport n'a pas été retrouvé dans la base de données.";
        }

        public string DownloadReport(int id)
        {
            SpaceReport report = db.SpaceReports.Find(id);
            if (report == null)
            {
                return HttpNotFound().ToString();
            }
            return Reports_Controller.Download(report.Id);
        }

        public string DeleteSpaceReport(int id)
        {
            try
            {
                SpaceReport report = db.SpaceReports.Find(id);
                Email email = report.Email;

                List<SpaceServer_Report> spaceserverreports = db.SpaceServerReports.ToList();
                foreach (SpaceServer_Report spaceserverreport in spaceserverreports)
                {
                    if (spaceserverreport.SpaceReport == report)
                    {
                        db.SpaceServerReports.Remove(spaceserverreport);
                        db.SaveChanges();
                    }
                }
                Specific_Logging(new Exception("...."), "DeleteSpaceReport", 3);
                return Reports_Controller.Delete(report.Id);
            }
            catch (Exception exception)
            {
                Specific_Logging(exception, "DeleteSpaceReport");
                return "Une erreur est surveunue lors de la suppression" +
                    exception.Message;
            }
        }

        public string Purge()
        {
            string message = "";
            List<SpaceReport> reports = db.SpaceReports.Where(rep => rep.Duration == null || rep.ResultPath == null).ToList();
            foreach (SpaceReport report in reports)
            {
                message += "Rapport " + report.DateTime + " supprimé";
                Email email = (report.Email != null) ? report.Email : null;
                List<SpaceServer_Report> spaceserverreports = report.SpaceServer_Reports.ToList();
                foreach (SpaceServer_Report spaceserverreport in spaceserverreports)
                {
                    db.SpaceServerReports.Remove(spaceserverreport);
                }
                db.SaveChanges();
                Reports_Controller.Delete(report.Id);
            }
            Specific_Logging(new Exception("...."), "Purge", 3);
            return message;
        }

        public bool IsAlreadyAdded(string servername)
        {
            int number = 0;
            if (servername != null && servername.Trim() != "")
            {
                number = db.SpaceServers
                .Where(ser => ser.Name.ToUpper() == servername.Trim().ToUpper()).Count();
            }
            return (number > 0);
        }

        public bool IsShare(string servername)
        {
            if (servername != null && servername.Trim() != "")
            {
                return servername.StartsWith(@"\\");
            }
            return false;
        }

        [HttpPost]
        public string AddSpaceServer()
        {
            string servername = "", partitions = "";
            string server_disks = "", include = "";
            try
            {
                servername = Request.Form["servername"].ToString();
                partitions = Request.Form["partitions"].ToString();
                string cellcolor = Request.Form["cellcolor"];
                include = Request.Form["include"].ToString();
                if (servername == null || servername.Trim() == ""
                    || IsAlreadyAdded(servername))
                {
                    return "Nom de serveur vide ou existant déjà.";
                }
                string[] disks = partitions.Split(new string[] { "; " }, StringSplitOptions.RemoveEmptyEntries);
                NumberFormatInfo provider = new NumberFormatInfo();
                provider.NumberDecimalSeparator = ".";
                foreach (string disk in disks)
                {
                    if (disk.Trim() == "")
                    {
                        continue;
                    }
                    string[] infos = disk.Split(new string[] { " || " }, StringSplitOptions.RemoveEmptyEntries);
                    if (infos.Length == 2)
                    {
                        string partition = infos[0].Trim();
                        double threshold = (Double)HomeController.SPACE_DEFAULT_THRESHOLD;
                        threshold = Math.Round(Convert.ToDouble(infos[1], provider), 2);
                        server_disks += "Disk=" + partition + "|Thr=" + threshold + " " +
                            HomeController.DEFAULT_TERA_OCTECT_UNIT + "; ";
                    }
                }
                if (server_disks.Length > 2)
                {
                    server_disks = server_disks.Substring(0, server_disks.Length - 2);
                }
                Dictionary<int, ServersController.VirtualizedServer> FOREST = ServersController.GetInformationsFromForestDomains();
                ReftechServers[] REFTECH_SERVERS = null;
                try
                {
                    REFTECH_SERVERS = db.ReftechServers.ToArray();
                }
                catch { }
                SpaceServer server = db.SpaceServers.Create();
                server.Name = servername.ToUpper();
                ServersController.VirtualizedServer_Result virtual_server = ServersController.GetServerInformations(FOREST, REFTECH_SERVERS, server, HomeController.SPACE_MODULE, false);
                server = virtual_server.SPACE_Server;
                server.Disks = server_disks;
                server.IsShare = IsShare(server.Name);
                server.Included = (include == "true");
                if (cellcolor != null)
                {
                    server.CellColor = "#" + cellcolor;
                }
                else
                {
                    server.CellColor = "#22ff22";
                }
                if (ModelState.IsValid)
                {
                    db.SpaceServers.Add(server);
                    db.SaveChanges();
                    Specific_Logging(new Exception("...."), "AddSpaceServer " + server.Name, 3);
                    return "Le serveur a été correctement rajouté";
                }
                Specific_Logging(new Exception("...."), "AddSpaceServer " + server.Name, 2);
                return "KO";
            }
            catch (Exception exception)
            {
                Specific_Logging(exception, "AddSpaceServer");
                return "Erreur lors de l'ajout du serveur";
            }
        }

        [HttpPost]
        public string EditSpaceServer(int id)
        {
            SpaceServer server = db.SpaceServers.Find(id);
            if (server == null)
            {
                return "Serveur non trouvé";
            }
            string servername = "", partitions = "";
            string server_disks = "";
            try
            {
                servername = Request.Form["servername"].ToString();
                partitions = Request.Form["partitions"].ToString();
                if ((servername == null || servername.Trim() == ""
                    || IsAlreadyAdded(servername)) && servername.Trim().ToUpper() != server.Name)
                {
                    return "Nom de serveur vide ou existant déjà.";
                }
                NumberFormatInfo provider = new NumberFormatInfo();
                provider.NumberDecimalSeparator = ".";
                string[] disks = partitions.Split(new string[] { "; " }, StringSplitOptions.RemoveEmptyEntries);
                foreach (string disk in disks)
                {
                    if (disk.Trim() == "")
                    {
                        continue;
                    }
                    string[] infos = disk.Split(new string[] { " || " }, StringSplitOptions.RemoveEmptyEntries);
                    if (infos.Length == 2)
                    {
                        string partition = infos[0].Trim();
                        double threshold = HomeController.SPACE_DEFAULT_THRESHOLD;
                        threshold = Math.Round(Convert.ToDouble(infos[1], provider), 2);
                        server_disks += "Disk=" + partition + "|Thr=" + threshold + " " +
                            HomeController.DEFAULT_TERA_OCTECT_UNIT + "; ";
                    }
                }
                if (server_disks.Length > 2)
                {
                    server_disks = server_disks.Substring(0, server_disks.Length - 2);
                }
                if (server.Name != servername.ToUpper())
                {
                    Dictionary<int, ServersController.VirtualizedServer> FOREST = ServersController.GetInformationsFromForestDomains();
                    ReftechServers[] REFTECH_SERVERS = null;
                    try
                    {
                        REFTECH_SERVERS = db.ReftechServers.ToArray();
                    }
                    catch { }
                    server.Name = servername.ToUpper();
                    ServersController.VirtualizedServer_Result virtual_server = ServersController.GetServerInformations(FOREST, REFTECH_SERVERS, server, HomeController.SPACE_MODULE, false);
                    server = virtual_server.SPACE_Server;
                    server.IsShare = IsShare(server.Name);
                }
                server.Disks = server_disks;
                if (ModelState.IsValid)
                {
                    db.Entry(server).State = System.Data.Entity.EntityState.Modified;
                    db.SaveChanges();
                    Specific_Logging(new Exception("...."), "EditSpaceServer " + server.Name, 3);
                    return "Le serveur a été correctement modifié";
                }
                Specific_Logging(new Exception("...."), "EditSpaceServer " + server.Name, 2);
                return "KO";
            }
            catch (Exception exception)
            {
                Specific_Logging(exception, "EditSpaceServer");
                return "Erreur lors de la modification du serveur";
            }
        }

        public string EditSpaceServerAccounts(int id)
        {
            try
            {
                string author = User.Identity.Name;
                SpaceServer server = db.SpaceServers.Find(id);
                string execution_account = Request.Form["execution_account"];
                string check_account = Request.Form["check_account"];
                string session_password = Request.Form["session_password"];
                if (!McoUtilities.IsValidLoginPassword(author, McoUtilities.Encrypt(session_password)))
                {
                    return "Authentification échouée:\n mauvaise combinaison username/password";
                }
                server.CheckAccount = check_account.ToUpper();
                server.ExecutionAccount = execution_account.ToUpper();
                string logs = "Check with " + server.CheckAccount;
                logs += "Exec with " + server.ExecutionAccount;
                if (check_account != null && check_account.Trim() != ""
                    && execution_account != null && execution_account.Trim() != "")
                {
                    if (ModelState.IsValid)
                    {
                        db.Entry(server).State = System.Data.Entity.EntityState.Modified;
                        db.SaveChanges();
                        Specific_Logging(new Exception("...."), "EditSpaceServerAccounts " + server.Name + " " + logs, 2);
                        return "Les modifications ont été effectuées sur le serveur.\n";
                    }
                }
                else
                {
                    Specific_Logging(new Exception("...."), "EditSpaceServerAccounts " + server.Name, 2);
                }
                return "Erreur lors de la modification des comptes du serveur";
            }
            catch (Exception exception)
            {
                Specific_Logging(exception, "EditSpaceServerAccounts");
                return "Erreur lors de la modification des comptes du serveur";
            }
        }

        public string GetAccountsList(int id)
        {
            SpaceServer server = db.SpaceServers.Find(id);
            if (server == null)
            {
                return Accounts_Controller.GetStringedAccountsList();
            }
            string options = "";
            List<Account> accounts = Accounts_Controller.GetAccountsList();
            foreach (Account account in accounts)
            {
                options += "<option val='" + account.DisplayName.ToUpper() + "'";
                if (server.CheckAccount != null && account.DisplayName.ToUpper() == server.CheckAccount.ToUpper())
                {
                    options += " selected";
                }
                options += ">" + account.DisplayName.ToUpper() + "</option>";
            }
            return options;
        }

        public string DeleteSpaceServer(int id)
        {
            SpaceServer server = db.SpaceServers.Find(id);
            string log = server.Name;
            SpaceServer_Report[] serverReports = server.SpaceServer_Reports.ToArray();
            foreach (SpaceServer_Report serverReport in serverReports)
            {
                db.SpaceServerReports.Remove(serverReport);
            }
            db.SpaceServers.Remove(server);
            db.SaveChanges();
            Specific_Logging(new Exception("...."), "DeleteSpaceServer " + server.Name, 3);
            return "Le serveur " + server.Name + " a été supprimé";
        }

        Color ContrastColor(string htmlcolor)
        {
            Color color = ColorTranslator.FromHtml(htmlcolor);
            int d = 0;
            // Counting the perceptive luminance - human eye favors green color... 
            double a = 1 - (0.299 * color.R + 0.587 * color.G + 0.114 * color.B) / 255;
            if (a < 0.5)
                d = 0; // bright colors - black font
            else
                d = 255; // dark colors - white font
            return Color.FromArgb(d, d, d);
        }

        public string DownloadInitFile()
        {
            try
            {
                System.Web.HttpResponse response = System.Web.HttpContext.Current.Response;
                response.ClearContent();
                response.Clear();
                response.ContentType = "text/plain";
                string filepath = HomeController.SPACE_RELATIVE_INIT_FILE;
                response.AddHeader("Content-Disposition", "attachment; filename=" + filepath + ";");
                String RelativePath = HomeController.SPACE_DEFAULT_INIT_FILE.Replace(Request.ServerVariables["APPL_PHYSICAL_PATH"], String.Empty);
                response.TransmitFile(HomeController.SPACE_DEFAULT_INIT_FILE);
                response.Flush();
                response.End();
                Specific_Logging(new Exception("...."), "DownloadInitFile", 3);
                return "OK";
            }
            catch (Exception exception)
            {
                Specific_Logging(exception, "DownloadInitFile");
                return exception.Message;
            }
        }

        public List<ServersController.VirtualizedPartition>
            GetPartitionsInfos(List<ServersController.VirtualizedPartition> partitions,
            Dictionary<string, string> server_disks)
        {
            List<ServersController.VirtualizedPartition> results
                = new List<ServersController.VirtualizedPartition>();
            foreach (KeyValuePair<string, string> disk in server_disks)
            {
                ServersController.VirtualizedPartition partition =
                    ServersController.GetVirtualizedPartition(partitions, disk.Key);
                if (partition != null)
                {
                    NumberFormatInfo provider = new NumberFormatInfo();
                    provider.NumberDecimalSeparator = HomeController.DEFAULT_NUMBER_DECIMAL_SEPARATOR;
                    string from = disk.Value.Split(' ')[1].Trim();
                    string value = disk.Value.Split(' ')[0].Trim();
                    double threshold = Math.Round(Convert.ToDouble(value, provider), 2);
                    partition.Threshold = ServersController.VirtualizedPartition.
                        SizeConversion(threshold, from, HomeController.DEFAULT_OCTECT_UNIT)
                        .ToString() + " " + HomeController.DEFAULT_OCTECT_UNIT;
                    bool test = partition.IsCritical();
                    results.Add(partition);
                }
            }
            return results;
        }

        public JsonResult CheckCapacityPlanning()
        {
            Dictionary<string, string> results = new Dictionary<string, string>();
            results["response"] = "";
            results["email"] = "";
            results["errors"] = "";
            int emailId = 0;

            SpaceReport report = db.SpaceReports.Create();
            report.DateTime = DateTime.Now;
            report.TotalChecked = 0;
            report.TotalErrors = 0;
            report.Module = HomeController.SPACE_MODULE;

            report.Author = User.Identity.Name;
            report.ResultPath = "";
            Email email = db.Emails.Create();
            email.Module = HomeController.SPACE_MODULE;
            email = Emails_Controller.SetRecipients(email, HomeController.SPACE_MODULE);
            report.Email = email;
            email.Report = report;
            if (ModelState.IsValid)
            {
                db.SpaceReports.Add(report);
                db.SaveChanges();
                emailId = report.Email.Id;
                int reportNumber = db.SpaceReports.Count();
                if (reportNumber > HomeController.SPACE_MAX_REPORT_NUMBER)
                {
                    int reportNumberToDelete = reportNumber - HomeController.SPACE_MAX_REPORT_NUMBER;
                    SpaceReport[] reportsToDelete =
                        db.SpaceReports.OrderBy(idReport => idReport.Id).Take(reportNumberToDelete).ToArray();
                    foreach (SpaceReport toDeleteReport in reportsToDelete)
                    {
                        DeleteSpaceReport(toDeleteReport.Id);
                    }
                }
            }
            DateTime Today = DateTime.Today;
            string FileName = "Check Capacity Planning " + DateTime.Now.ToString("dd") + "_" + DateTime.Now.ToString("MM") + "_" +
                DateTime.Now.ToString("yyyy") + report.Id + ".xlsx";

            List<SpaceServer> servers = new List<SpaceServer>();
            try
            {
                servers = db.SpaceServers.Where(ser => ser.Included == true).ToList();
            }
            catch { }
            foreach (SpaceServer server in servers)
            {
                Dictionary<string, string> server_disks = new Dictionary<string, string>();
                string[] parser = server.Disks.Split(new string[] { "; " }, StringSplitOptions.RemoveEmptyEntries);
                for (int index = 0; index < parser.Length; index++)
                {
                    string[] infos = parser[index].Split(new string[] { "|" }, StringSplitOptions.RemoveEmptyEntries);
                    string disk = infos[0].Substring(infos[0].IndexOf("=") + 1);
                    string threshold = infos[1].Substring(infos[1].IndexOf("=") + 1);
                    server_disks.Add(disk, threshold);
                }

                SpaceServer_Report serverreport = db.SpaceServerReports.Create();
                serverreport.Details = serverreport.Ping = serverreport.State = "";
                serverreport.SpaceReport = report;
                serverreport.SpaceReportId = report.Id;
                serverreport.SpaceServer = server;
                serverreport.SpaceServerId = server.Id;
                List<ServersController.VirtualizedPartition> partitions;
                partitions = GetRemainingSpace(server);
                partitions = GetPartitionsInfos(partitions, server_disks);
                List<bool> criticals = new List<bool>();
                foreach (ServersController.VirtualizedPartition partition in partitions)
                {
                    criticals.Add(partition.Critical);
                    serverreport.Details += "Disk=" + partition.Name + "|Free=" +
                            partition.AvailableSpace;
                    if (partition.Critical)
                    {
                        serverreport.Details += "|KO";
                    }
                    else
                    {
                        serverreport.Details += "|OK";
                    }
                    serverreport.Details += "; ";
                }
                if (serverreport.Details.Length > 2)
                {
                    serverreport.Details = serverreport.Details.Substring(0, serverreport.Details.Length - 2);
                }
                serverreport.State = (!criticals.Contains(false)) ? "OK" :
                    (!criticals.Contains(true)) ? "KO" : "H-OK";
                if (ModelState.IsValid)
                {
                    db.SpaceServerReports.Add(serverreport);
                    db.SaveChanges();
                }
            }
            report.Duration = DateTime.Now.Subtract(report.DateTime);
            report.TotalChecked = report.SpaceServer_Reports.Count;
            report.TotalErrors = report.SpaceServer_Reports.Where(ser => ser.State != "OK").Count();
            if (ModelState.IsValid)
            {
                db.Entry(report).State = System.Data.Entity.EntityState.Modified;
                db.SaveChanges();
            }
            report = RegisterResults(report);
            if (ModelState.IsValid)
            {
                db.Entry(report).State = System.Data.Entity.EntityState.Modified;
                db.SaveChanges();
            }
            BuildEmail(emailId);
            results["response"] = "OK";
            results["email"] = emailId.ToString();
            results["errors"] = "Fin d'exécution. \n" + "Erreurs d'exécution : " + results["errors"];
            Specific_Logging(new Exception("...."), "CheckCapacityPlanning", 3);
            return Json(results, JsonRequestBehavior.AllowGet);
        }

        public bool EmptyCharts()
        {
            try
            {
                string sheetName = "CapacityGraph";

                bool found_sheet = false;
                foreach (Excel.Worksheet sheet in MyWorkbook.Sheets)
                {
                    if (sheet.Name.StartsWith(sheetName))
                    {
                        sheetName = sheet.Name;
                        found_sheet = true;
                        break;
                    }
                }
                if (found_sheet)
                {
                    Excel._Worksheet GraphSheet = (Excel.Worksheet)MyWorkbook.Sheets[sheetName];
                    GraphSheet.Delete();
                }
            }
            catch { return false; }
            return true;

        }

        public bool DrawChart(SpaceServer server, Excel.Range cell, string disk, int top, int left)
        {
            long limit = cell.Row - HomeController.SPACE_MAX_CHARTS_LINES_NUMBER;
            int start = (limit < 0) ? 5 : (int)limit;

            string title = (server.IsShare) ? disk : server.Name + "Disque " + disk;
            try
            {
                string sheetName = "CapacityGraph";

                Excel._Worksheet DataSheet = (Excel.Worksheet)MyWorkbook.Sheets["Relevé SUP"];
                bool found_sheet = false;
                foreach (Excel.Worksheet sheet in MyWorkbook.Sheets)
                {
                    if (sheet.Name.StartsWith(sheetName))
                    {
                        sheetName = sheet.Name;
                        found_sheet = true;
                        break;
                    }
                }
                if (found_sheet)
                {
                    MySheet = (Excel.Worksheet)MyWorkbook.Sheets[sheetName];
                }
                else
                {
                    MySheet = (Excel.Worksheet)MyWorkbook.Sheets.Add();
                    MySheet.Name = sheetName;
                }
                MySheet.Activate();

                // Add chart.
                string chart_name = (!server.IsShare) ? server.Name + " Disque " + disk : disk;
                Excel.ChartObjects charts = MySheet.ChartObjects();
                Excel.ChartObject chartObject = null;
                foreach (Excel.ChartObject chart_obj in charts)
                {
                    if (chart_obj.Name == chart_name)
                    {
                        chartObject = chart_obj;
                        break;
                    }
                }
                if (chartObject == null)
                {
                    chartObject = charts.Add(10 + left, 10 + top, 300, 300);
                    chartObject.Name = chart_name;
                }
                Excel.Chart chart = chartObject.Chart;
                chart.Legend.Position = Excel.XlLegendPosition.xlLegendPositionTop;
                //chart.Legend.Height = 10;
                // Set chart range.
                Excel.Range date_range = DataSheet.Range[DataSheet.Cells[start, 1], DataSheet.Cells[cell.Row, 1]];
                //chart.SetSourceData(range);

                Excel.SeriesCollection seriesCollection;
                Excel.Series threshold, spaces;

                //this create the seriescollection and series
                seriesCollection = chart.SeriesCollection();
                threshold = seriesCollection.NewSeries();
                spaces = seriesCollection.NewSeries();


                //this gives each series the values. values are y values, xvalues are x values 
                //threshold.XValues = date_range;
                threshold.Values = DataSheet.Range[DataSheet.Cells[start, cell.Column + 1], DataSheet.Cells[cell.Row, cell.Column + 1]];
                threshold.Name = "Seuil";

                spaces.XValues = date_range;
                spaces.Values = DataSheet.Range[DataSheet.Cells[start, cell.Column], DataSheet.Cells[cell.Row, cell.Column]];
                spaces.Name = chart_name;
                //spaces.Format.Line.ForeColor.RGB = ColorTranslator.FromHtml(server.CellColor).ToArgb();
                // Set chart properties.
                chart.ChartType = Excel.XlChartType.xlLine;
                chart.PlotVisibleOnly = false;

            }
            catch (Exception exception)
            {
                Specific_Logging(exception, "DrawChart");
                string message = exception.Message;
            }
            MySheet = (Excel.Worksheet)MyWorkbook.Sheets["Relevé SUP"];
            MySheet.Activate();
            return true;
        }

        public bool InitFile()
        {
            try
            {
                NumberFormatInfo provider = new NumberFormatInfo();
                provider.NumberDecimalSeparator = HomeController.DEFAULT_NUMBER_DECIMAL_SEPARATOR;
                ICollection<SpaceServer> list = db.SpaceServers.OrderBy(ser => ser.Name).ToArray();
                string filename = "";
                MySheet.Cells[1, 1] = "Serveurs";
                MySheet.Cells[2, 1] = "Partitions";

                MySheet.Cells[4, 1] = "Seuils";
                MySheet.Cells[4, 1].EntireRow.Font.Bold = true;
                MySheet.Cells[1, 1].EntireRow.Font.Bold = true;
                MySheet.Cells[1, 1].EntireColumn.Font.Bold = true;
                MySheet.Cells[1, 1].ColumnWidth = 20;
                MySheet.Cells.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                MySheet.Cells.Style.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                MySheet.Cells.Style.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                MySheet.Cells[2, 1].EntireRow.Font.Bold = true;
                int column_index = 2;
                foreach (SpaceServer server in list)
                {
                    int start_column_index = column_index;
                    Dictionary<string, string> server_disks = new Dictionary<string, string>();
                    string[] parser = server.Disks.Split(new string[] { "; " }, StringSplitOptions.RemoveEmptyEntries);
                    for (int index = 0; index < parser.Length; index++)
                    {
                        string[] infos = parser[index].Split(new string[] { "|" }, StringSplitOptions.RemoveEmptyEntries);
                        string disk = infos[0].Substring(infos[0].IndexOf("=") + 1).Trim();
                        string threshold = infos[1].Substring(infos[1].IndexOf("=") + 1).Trim();
                        threshold = threshold.Split(' ')[0].Trim();
                        server_disks.Add(disk, threshold);
                    }
                    MySheet.Cells[1, column_index + 1] = "Seuil";
                    int number_to_merge = 0;
                    MySheet.Cells[1, column_index] = server.Name;
                    foreach (KeyValuePair<string, string> disk in server_disks)
                    {
                        MySheet.Cells[2, column_index] = (server.IsShare) ?
                            disk.Key : "Disque " + disk.Key;
                        MySheet.Cells[4, column_index] = Convert.ToDouble(disk.Value, provider);
                        MySheet.Cells[4, column_index].NumberFormat = "# ##0,00\" To\"";
                        MySheet.Cells[4, column_index + 1] = Convert.ToDouble(disk.Value, provider);
                        MySheet.Cells[4, column_index + 1].EntireColumn.Hidden = true;
                        column_index += 2;
                        number_to_merge += 2;
                    }
                    Color rangeColor = ContrastColor(server.CellColor);
                    number_to_merge = number_to_merge - 1;
                    Excel.Range to_merge = MySheet.Range[MySheet.Cells[1, start_column_index], MySheet.Cells[1, start_column_index + number_to_merge]];
                    to_merge.Merge();
                    to_merge.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                    to_merge.EntireColumn.Interior.Color = System.Drawing.ColorTranslator.FromHtml(server.CellColor);
                    to_merge.EntireColumn.Font.Color = System.Drawing.ColorTranslator
                        .FromHtml(ColorTranslator.ToHtml(ContrastColor(server.CellColor)));
                    to_merge.Font.Bold = true;
                    to_merge.Style.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                    to_merge.Style.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                }
                MySheet.Cells[4, 1].EntireRow.Interior.Color = System.Drawing.ColorTranslator.FromHtml("#ffc000");
                MySheet.Cells[4, 1].EntireRow.Font.Color = System.Drawing.ColorTranslator.FromHtml("#000");
                filename = "Controle_Capacity_Planning_" +
                    DateTime.Now.ToString("MM") + "_" + DateTime.Now.ToString("yyyy") + ".xlsx";
                /*MyWorkbook.SaveAs(HomeController.SPACE_RESULTS_FOLDER + filename,
                        Excel.XlFileFormat.xlOpenXMLWorkbook, null, null, null, null,
                        Excel.XlSaveAsAccessMode.xlNoChange, null, null, null, null, null);*/
            }
            catch (Exception exception)
            {
                Specific_Logging(exception, "InitFile");
                return false;
            }
            Specific_Logging(new Exception("...."), "InitFile", 3);
            return true;
        }

        public Excel.Range UpdateFile(SpaceServer server)
        {
            //NEW SERVER
            NumberFormatInfo provider = new NumberFormatInfo();
            provider.NumberDecimalSeparator = HomeController.DEFAULT_NUMBER_DECIMAL_SEPARATOR;

            int lastCol = MySheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Column + 1;
            if (MySheet.Cells[2, lastCol].value != null)
            {
                lastCol += 2;
            }
            if (lastCol > 1 && MySheet.Cells[2, lastCol - 1].value != null
                && MySheet.Cells[2, lastCol].value == null)
            {
                lastCol++;
            }

            int start_column_index = lastCol;
            Dictionary<string, string> server_disks = new Dictionary<string, string>();
            string[] serv_parser = server.Disks.Split(new string[] { "; " }, StringSplitOptions.RemoveEmptyEntries);
            for (int index = 0; index < serv_parser.Length; index++)
            {
                string[] infos = serv_parser[index].Split(new string[] { "|" }, StringSplitOptions.RemoveEmptyEntries);
                string disk = infos[0].Substring(infos[0].IndexOf("=") + 1).Trim();
                string threshold = infos[1].Substring(infos[1].IndexOf("=") + 1).Trim();
                server_disks.Add(disk, threshold);
            }

            int number_to_merge = 0;
            MySheet.Cells[1, lastCol] = server.Name;
            MySheet.Cells[1, lastCol + 1] = "Seuil";

            foreach (KeyValuePair<string, string> disk in server_disks)
            {
                string value = disk.Value.Split(' ')[0];
                string partition = (disk.Key.StartsWith(@"\\")) ? disk.Key
                    : "Disque " + disk.Key;
                MySheet.Cells[2, lastCol] = partition;
                MySheet.Cells[4, lastCol] = Convert.ToDouble(value, provider);
                MySheet.Cells[4, lastCol].NumberFormat = "# ##0,00\" To\"";
                MySheet.Cells[4, lastCol].EntireColumn.ColumnWidth = 15;
                MySheet.Cells[4, lastCol + 1] = Convert.ToDouble(value, provider);
                MySheet.Cells[4, lastCol + 1].EntireColumn.Hidden = true;

                lastCol += 2;
                number_to_merge += 2;
            }
            Color rangeColor = ContrastColor(server.CellColor);
            number_to_merge = number_to_merge - 1;
            Excel.Range to_merge = MySheet.Range[MySheet.Cells[1, start_column_index], MySheet.Cells[1, start_column_index + number_to_merge]];
            to_merge.Merge();
            to_merge.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
            to_merge.EntireColumn.Interior.Color = System.Drawing.ColorTranslator.FromHtml(server.CellColor);
            to_merge.EntireColumn.Font.Color = System.Drawing.ColorTranslator
                .FromHtml(System.Drawing.ColorTranslator.ToHtml(ContrastColor(server.CellColor)));
            to_merge.Font.Bold = true;
            to_merge.Style.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
            to_merge.Style.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

            MySheet.Cells[4, 1].EntireRow.Interior.Color = System.Drawing.ColorTranslator.FromHtml("#ffc000");
            MySheet.Cells[4, 1].EntireRow.Font.Color = System.Drawing.ColorTranslator.FromHtml("#000");
            Specific_Logging(new Exception("...."), "UpdateFile " + server.Name, 3);
            return to_merge;
        }

        public SpaceReport RegisterResults(SpaceReport report)
        {
            string Errors = "";
            try
            {
                NumberFormatInfo provider = new NumberFormatInfo();
                provider.NumberDecimalSeparator = HomeController.DEFAULT_NUMBER_DECIMAL_SEPARATOR;
                MyApplication = new Excel.Application();
                MyApplication.Visible = false;
                MyApplication.DisplayAlerts = false;

                bool found = false;
                string filename = "";
                string[] report_files = System.IO.Directory.GetFiles(HomeController.SPACE_RESULTS_FOLDER,
                    "Controle_Capacity_Planning_*.xlsx", System.IO.SearchOption.TopDirectoryOnly);

                if (report_files.Length > 0)
                {
                    filename = Directory.GetFiles(HomeController.SPACE_RESULTS_FOLDER,
                        "Controle_Capacity_Planning_*.xlsx", System.IO.SearchOption.TopDirectoryOnly)
                            .Select(x => new FileInfo(x))
                            .OrderByDescending(x => x.LastWriteTime).FirstOrDefault().FullName;
                    MyWorkbook = MyApplication.Workbooks.Open(filename);
                    found = true;
                }
                else
                {
                    Errors += HomeController.SPACE_RESULTS_FOLDER + "Controle_hebdomadaire_sauvegarde_BESR_PCI_(Date)-(id).xlsx n'a pas été trouvé.\r\n";
                    MyWorkbook = MyApplication.Workbooks.Add(Excel.XlWBATemplate.xlWBATWorksheet);
                    found = false;
                }

                //SHEET 
                string sheetName = "Relevé SUP";
                bool found_sheet = false;
                foreach (Excel.Worksheet sheet in MyWorkbook.Sheets)
                {
                    if (sheet.Name.StartsWith(sheetName))
                    {
                        sheetName = sheet.Name;
                        found_sheet = true;
                        break;
                    }
                }
                if (found_sheet)
                {
                    MySheet = (Excel.Worksheet)MyWorkbook.Sheets[sheetName];
                }
                else
                {
                    MySheet = (Excel.Worksheet)MyWorkbook.Sheets.Add();
                    MySheet.Name = sheetName;
                }
                MySheet.Activate();

                //MANAGE 
                //
                ICollection<SpaceServer> list = db.SpaceServers.OrderBy(ser => ser.Name).ToArray();
                if (!found_sheet)
                {
                    InitFile();
                }
                int lastRow = MySheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row + 1;
                MySheet.Cells[lastRow, 1] = report.DateTime.ToString("dd") +
                    "/" + report.DateTime.ToString("MM") + "/" +
                    report.DateTime.ToString("yyyy") + " " + report.DateTime.ToString("HH") + " : " + report.DateTime.ToString("mm");
                List<SpaceServer_Report> serverreports = report.SpaceServer_Reports.ToList();
                Excel.Range server_searcher = MySheet.Range[MySheet.Cells[1, 2], MySheet.Cells[1, 100]];
                int chart_index = 0;
                int chart_top = 0, chart_left = 0;
                EmptyCharts();
                foreach (SpaceServer_Report serverreport in serverreports)
                {
                    if (chart_index > 3)
                    {
                        chart_left = 0;
                        chart_index = 0;
                        chart_top += 305;
                    }
                    string[] parser = serverreport.Details.Split(new string[] { "; " }, StringSplitOptions.RemoveEmptyEntries);
                    bool found_server = false;
                    Excel.Range cell_server = null;
                    foreach (Excel.Range cell in server_searcher)
                    {
                        if (cell.Value != null && cell.Value == serverreport.SpaceServer.Name)
                        {
                            found_server = true;
                            cell_server = cell;
                            break;
                        }
                    }

                    if (!found_server)
                    {
                        cell_server = UpdateFile(serverreport.SpaceServer);
                    }
                    Excel.Range disk_searcher = MySheet.Range[MySheet.Cells[2, cell_server.Column], MySheet.Cells[2, cell_server.Column + 10]];
                    for (int index = 0; index < parser.Length; index++)
                    {
                        string[] infos = parser[index].Split(new string[] { "|" }, StringSplitOptions.RemoveEmptyEntries);
                        string disk = infos[0].Substring(infos[0].IndexOf("=") + 1).Trim();
                        string free = infos[1].Substring(infos[1].IndexOf("=") + 1).Trim();
                        free = free.Split(' ')[0].Trim();
                        string state = infos[2].Trim();
                        foreach (Excel.Range col in disk_searcher)
                        {
                            if (col.Value != null
                                && (col.Value.Trim() == "Disque " + disk || col.Value.ToLower() == disk.ToLower())
                               )
                            {
                                MySheet.Cells[lastRow, col.Column] = Convert.ToDouble(free, provider);
                                MySheet.Cells[lastRow, col.Column].NumberFormat = "# ##0,00\" To\"";
                                MySheet.Cells[lastRow, col.Column].EntireColumn.ColumnWidth = 15;
                                MySheet.Cells[lastRow, col.Column + 1] = MySheet.Cells[4, col.Column].Value;
                                if (state != "OK")
                                {
                                    MySheet.Cells[lastRow, col.Column].Font.Color =
                                        System.Drawing.ColorTranslator.FromHtml("#ff2f00");
                                    MySheet.Cells[lastRow, col.Column].Font.Bold = true;
                                }
                                DrawChart(serverreport.SpaceServer, MySheet.Cells[lastRow, col.Column], disk, chart_top, chart_left);
                                chart_left += 305;
                                chart_index++;

                                break;
                            }
                        }
                    }
                }
                filename = "Controle_Capacity_Planning_" +
                    DateTime.Now.ToString("MM") + "_" + DateTime.Now.ToString("yyyy") + ".xlsx";
                MyWorkbook.SaveAs(HomeController.SPACE_RESULTS_FOLDER + filename,
                            Excel.XlFileFormat.xlOpenXMLWorkbook, null, null, null, null,
                            Excel.XlSaveAsAccessMode.xlNoChange, null, null, null, null, null);
                report.ResultPath = HomeController.SPACE_RESULTS_FOLDER + "Controle_Capacity_Planning_" +
                    DateTime.Now.ToString("dd") + "_" + DateTime.Now.ToString("MM") + "_" +
                    DateTime.Now.ToString("yyyy") + "_ID_" + report.Id + ".xlsx";
                MyWorkbook.SaveCopyAs(report.ResultPath);
            }
            catch (Exception exception)
            {
                Specific_Logging(exception, "RegisterResults " + report.Id);
                string message = exception.Message;
            }
            finally
            {
                McoUtilities.CloseExcel(MyApplication, MyWorkbook, MySheet);
            }
            Specific_Logging(new Exception("...."), "RegisterResults " + report.Id, 3);
            return report;
        }

        public bool RegisterReportInfos(SpaceReport report, int lastRow)
        {
            NumberFormatInfo provider = new NumberFormatInfo();
            provider.NumberDecimalSeparator = HomeController.DEFAULT_NUMBER_DECIMAL_SEPARATOR;
            List<SpaceServer_Report> serverreports = report.SpaceServer_Reports.ToList();
            Excel.Range server_searcher = MySheet.Range[MySheet.Cells[1, 2], MySheet.Cells[1, 100]];
            int chart_index = 0;
            int chart_top = 0, chart_left = 0;
            EmptyCharts();
            foreach (SpaceServer_Report serverreport in serverreports)
            {
                if (chart_index > 3)
                {
                    chart_left = 0;
                    chart_index = 0;
                    chart_top += 305;
                }
                string[] parser = serverreport.Details.Split(new string[] { "; " }, StringSplitOptions.RemoveEmptyEntries);
                bool found_server = false;
                Excel.Range cell_server = null;
                foreach (Excel.Range cell in server_searcher)
                {
                    if (cell.Value != null && cell.Value == serverreport.SpaceServer.Name)
                    {
                        found_server = true;
                        cell_server = cell;
                        break;
                    }
                }

                if (!found_server)
                {
                    cell_server = UpdateFile(serverreport.SpaceServer);
                }
                Excel.Range disk_searcher = MySheet.Range[MySheet.Cells[2, cell_server.Column], MySheet.Cells[2, cell_server.Column + 10]];
                for (int index = 0; index < parser.Length; index++)
                {
                    string[] infos = parser[index].Split(new string[] { "|" }, StringSplitOptions.RemoveEmptyEntries);
                    string disk = infos[0].Substring(infos[0].IndexOf("=") + 1).Trim();
                    string free = infos[1].Substring(infos[1].IndexOf("=") + 1).Trim();
                    free = free.Split(' ')[0].Trim();
                    string state = infos[2];
                    foreach (Excel.Range col in disk_searcher)
                    {
                        if (col.Value != null
                            && (col.Value == "Disque " + disk || col.Value.ToLower() == disk.ToLower())
                           )
                        {
                            MySheet.Cells[lastRow, col.Column] = Convert.ToDouble(free, provider);
                            MySheet.Cells[lastRow, col.Column].NumberFormat = "# ##0,00\" To\"";
                            MySheet.Cells[lastRow, col.Column].EntireColumn.ColumnWidth = 15;
                            MySheet.Cells[lastRow, col.Column + 1] = MySheet.Cells[4, col.Column].Value;
                            if (state != "OK")
                            {
                                MySheet.Cells[lastRow, col.Column].Font.Color =
                                    System.Drawing.ColorTranslator.FromHtml("#ff2f00");
                                MySheet.Cells[lastRow, col.Column].Font.Bold = true;
                            }
                            DrawChart(serverreport.SpaceServer, MySheet.Cells[lastRow, col.Column], disk, chart_top, chart_left);
                            chart_left += 305;
                            chart_index++;

                            break;
                        }
                    }
                }
            }
            return true;
        }

        public string Export()
        {
            string message = "", filename = "";
            try
            {
                NumberFormatInfo provider = new NumberFormatInfo();
                provider.NumberDecimalSeparator = HomeController.DEFAULT_NUMBER_DECIMAL_SEPARATOR;
                MyApplication = new Excel.Application();
                MyApplication.Visible = false;
                MyApplication.DisplayAlerts = false;

                bool found = false;
                MyWorkbook = MyApplication.Workbooks.Add(Excel.XlWBATemplate.xlWBATWorksheet);

                //SHEET 
                string sheetName = "Relevé SUP";
                MySheet = (Excel.Worksheet)MyWorkbook.Sheets.Add();
                MySheet.Name = sheetName;

                MySheet.Activate();
                InitFile();
                List<SpaceReport> reports = db.SpaceReports.OrderBy(rep => rep.DateTime).ToList();
                foreach (SpaceReport report in reports)
                {
                    int lastRow = MySheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row + 1;
                    MySheet.Cells[lastRow, 1] = report.DateTime.ToString("dd") +
                        "/" + report.DateTime.ToString("MM") + "/" +
                        report.DateTime.ToString("yyyy") + " " + report.DateTime.ToString("HH") + " : " + report.DateTime.ToString("mm");
                    List<SpaceServer_Report> serverreports = report.SpaceServer_Reports.ToList();
                    RegisterReportInfos(report, lastRow);
                }
                filename = HomeController.SPACE_RESULTS_FOLDER + "ExportCapacityPlanning" + DateTime.Now.ToString("dd") +
                        DateTime.Now.ToString("MM") + DateTime.Now.ToString("yyyy") + ".xlsx";
                try
                {
                    if (System.IO.File.Exists(filename))
                    {
                        System.IO.File.Delete(filename);
                    }
                }
                catch { }

                MyWorkbook.SaveAs(filename,
                            Excel.XlFileFormat.xlOpenXMLWorkbook, null, null, null, null,
                            Excel.XlSaveAsAccessMode.xlNoChange, null, null, null, null, null);
                message = "OK";
            }
            catch (Exception exception)
            {
                Specific_Logging(exception, "Export");
                string error = exception.Message;
                message = "KO";
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
                    response.AddHeader("Content-Disposition", "attachment; filename=" + filename + ";");
                    String RelativePath = filename.Replace(Request.ServerVariables["APPL_PHYSICAL_PATH"], String.Empty);
                    response.TransmitFile(filename);
                    response.Flush();
                    response.End();
                    Specific_Logging(new Exception("...."), "Export", 3);
                    return "OK";
                }
                catch (Exception exception)
                {
                    Specific_Logging(exception, "Export");
                    return exception.Message;
                }
            }
            Specific_Logging(new Exception("...."), "Export", 2);
            return message;
        }

        public Account GetSpaceServerAccount(int id = 0)
        {
            Account account = new Account();
            if (id != 0)
            {
                SpaceServer server = db.SpaceServers.Find(id);
                if (server != null)
                {
                    account = db.Accounts.Where(acc => acc.DisplayName.ToUpper() == server.CheckAccount.ToUpper())
                        .FirstOrDefault();
                }
            }
            if (account == null || account.DisplayName == null || account.DisplayName.Trim() == "")
            {
                account = new Account();
                account.DisplayName = HomeController.SPACE_DOMAIN_AND_USERNAME_IMPERSONNATION;
                account.Username = HomeController.SPACE_USERNAME_IMPERSONNATION;
                account.Domain = HomeController.DEFAULT_DOMAIN_IMPERSONNATION;
                account.Password = HomeController.SPACE_PASSWORD_IMPERSONNATION;
            }
            return account;
        }

        public List<ServersController.VirtualizedPartition> GetRemainingSpace(SpaceServer server)
        {
            List<ServersController.VirtualizedPartition> partitions =
                new List<ServersController.VirtualizedPartition>();
            Account account = Accounts_Controller.GetAccountByDisplayName(server.ExecutionAccount);
            if (account == null) { account = GetSpaceServerAccount(server.Id); }
            if (!server.IsShare)
            {
                partitions = ServersController.GetRemainingSpaceOnDisks(server.Name, account);
            }
            else
            {
                partitions = ServersController.GetRemainingSpaceOnMappedPartitions(server, account);
            }
            return partitions;
        }

        [HttpPost]
        public string Import(bool import)
        {
            string message = "";
            if (import)
            {
                SpaceReport[] reports = db.SpaceReports.ToArray();
                foreach (SpaceReport report in reports)
                {
                    message += DeleteSpaceReport(report.Id) + "\n <br />";
                }

                List<SpaceServer> servers = db.SpaceServers.ToList();
                foreach (SpaceServer server in servers)
                {
                    message += DeleteSpaceServer(server.Id) + "\n <br />";
                }
                servers = new List<SpaceServer>();
                string sheetName = "Relevé SUP";
                MyApplication = new Excel.Application();
                MyApplication.Visible = false;
                try
                {
                    MyWorkbook = MyApplication.Workbooks.Open(HomeController.SPACE_INIT_FILE);
                    MySheet = (Excel.Worksheet)MyWorkbook.Sheets[sheetName]; // Explicit cast is not required here
                    int lastCol = MySheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Column + 1;
                    Dictionary<int, ServersController.VirtualizedServer> FOREST = ServersController.GetInformationsFromForestDomains();
                    ReftechServers[] REFTECH_SERVERS = null;
                    try
                    {
                        REFTECH_SERVERS = db.ReftechServers.ToArray();
                    }
                    catch { }
                    Dictionary<string, int> servernames = new Dictionary<string, int>();
                    Excel.Range servers_searcher = MySheet.Range[MySheet.Cells[1, 2], MySheet.Cells[1, lastCol]];
                    foreach (Excel.Range cell in servers_searcher)
                    {
                        var content = cell.Value;
                        if (content != null)
                        {
                            servernames.Add(content, 0);
                            string disks = "";
                            Color color = ColorTranslator.FromOle((int)cell.Interior.Color);
                            string cellcolor = ColorTranslator.ToHtml(color);
                            Excel.Range disks_searcher = MySheet.Range[MySheet.Cells[2, cell.Column], MySheet.Cells[2, lastCol]];
                            Excel.Range thresholds_searcher = MySheet.Range[MySheet.Cells[4, cell.Column], MySheet.Cells[4, lastCol]];
                            foreach (Excel.Range col in disks_searcher)
                            {
                                if (col.Value == null && (col.Next == null || col.Next.Value == null))
                                {
                                    break;
                                }
                                if (col.Value == null)
                                {
                                    continue;
                                }
                                else
                                {
                                    string value = (string)col.Value;
                                    var test = MySheet.Cells[1, col.Column].Value;
                                    if (test != null && test != content)
                                    {
                                        goto ADD_SERVER;
                                    }
                                    value = (value.IndexOf("Disque ") != -1) ? value.Substring(value.IndexOf(" "))
                                        : value;
                                    disks += "Disk=" + value.ToUpper() + "|";
                                    double threshold = HomeController.SPACE_DEFAULT_THRESHOLD;
                                    foreach (Excel.Range val in thresholds_searcher)
                                    {
                                        if (col.Column == val.Column)
                                        {
                                            threshold = (double)val.Value;
                                            break;
                                        }
                                    }
                                    disks += threshold + " To; ";
                                }
                            }
                        ADD_SERVER:
                            {
                                if (disks.Length > 2)
                                {
                                    disks = disks.Substring(0, disks.Length - 2);
                                }
                            }
                            SpaceServer server = db.SpaceServers.Create();
                            server.Name = content;
                            server.Name = server.Name.ToUpper();
                            ServersController.VirtualizedServer_Result virtual_server = ServersController.GetServerInformations(FOREST, REFTECH_SERVERS, server, HomeController.SPACE_MODULE, false);
                            server = virtual_server.SPACE_Server;
                            server.IsShare = server.Name.StartsWith(@"\\");
                            server.Disks = disks;
                            server.CheckAccount = HomeController.SPACE_DOMAIN_AND_USERNAME_IMPERSONNATION;
                            server.ExecutionAccount = HomeController.SPACE_DOMAIN_AND_USERNAME_IMPERSONNATION;
                            server.Included = true;
                            server.CellColor = cellcolor;
                            servers.Add(server);
                        }
                    }
                    foreach (SpaceServer server in servers)
                    {
                        if (ModelState.IsValid)
                        {
                            db.SpaceServers.Add(server);
                            message += "Serveur " + server.Name + " Partitions:" + server.Disks + " Rajouté<br/>\n";
                        }
                    }
                    db.SaveChanges();
                }
                catch (Exception exception)
                {
                    Specific_Logging(exception, "Import");
                    message += "Des erreurs ont été signalées lors de la génération des Pools.";
                }
                finally
                {
                    McoUtilities.CloseExcel(MyApplication, MyWorkbook, MySheet);
                }
            }
            else
            {
                return "Erreur d'importation.";
            }
            Specific_Logging(new Exception("...."), "Import", 3);
            return message;
        }

        public JsonResult ExecuteSchedule(int id)
        {
            SpaceSchedule schedule = db.SpaceSchedules.Find(id);
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
            results["response"] = "";
            results["email"] = "";
            results["errors"] = "";
            int emailId = 0;

            SpaceReport report = db.SpaceReports.Create();
            report.DateTime = DateTime.Now;
            report.TotalChecked = 0;
            report.TotalErrors = 0;
            report.Module = HomeController.SPACE_MODULE;
            report.ScheduleId = schedule.Id;
            report.Schedule = schedule;
            report.Author = User.Identity.Name;
            report.ResultPath = "";
            Email email = db.Emails.Create();
            email.Module = HomeController.SPACE_MODULE;
            email = Emails_Controller.SetRecipients(email, HomeController.SPACE_MODULE);
            report.Email = email;
            email.Report = report;
            if (ModelState.IsValid)
            {
                db.SpaceReports.Add(report);
                db.SaveChanges();
                emailId = report.Email.Id;
                int reportNumber = db.SpaceReports.Count();
                if (reportNumber > HomeController.SPACE_MAX_REPORT_NUMBER)
                {
                    int reportNumberToDelete = reportNumber - HomeController.SPACE_MAX_REPORT_NUMBER;
                    SpaceReport[] reportsToDelete =
                        db.SpaceReports.OrderBy(idReport => idReport.Id).Take(reportNumberToDelete).ToArray();
                    foreach (SpaceReport toDeleteReport in reportsToDelete)
                    {
                        DeleteSpaceReport(toDeleteReport.Id);
                    }
                }
            }
            DateTime Today = DateTime.Today;
            string FileName = "Check Capacity Planning " + DateTime.Now.ToString("dd") + "_" + DateTime.Now.ToString("MM") + "_" +
                DateTime.Now.ToString("yyyy") + report.Id + ".xlsx";

            List<SpaceServer> servers = new List<SpaceServer>();
            try
            {
                servers = db.SpaceServers.Where(ser => ser.Included == true).ToList();
            }
            catch { }
            foreach (SpaceServer server in servers)
            {
                Dictionary<string, string> server_disks = new Dictionary<string, string>();
                string[] parser = server.Disks.Split(new string[] { "; " }, StringSplitOptions.RemoveEmptyEntries);
                for (int index = 0; index < parser.Length; index++)
                {
                    string[] infos = parser[index].Split(new string[] { "|" }, StringSplitOptions.RemoveEmptyEntries);
                    string disk = infos[0].Substring(infos[0].IndexOf("=") + 1);
                    string threshold = infos[1].Substring(infos[1].IndexOf("=") + 1);
                    server_disks.Add(disk, threshold);
                }

                SpaceServer_Report serverreport = db.SpaceServerReports.Create();
                serverreport.Details = serverreport.Ping = serverreport.State = "";
                serverreport.SpaceReport = report;
                serverreport.SpaceServer = server;
                serverreport.SpaceReportId = report.Id;
                serverreport.SpaceServerId = server.Id;
                List<ServersController.VirtualizedPartition> partitions;
                if (!server.IsShare)
                {
                    partitions = ServersController.GetRemainingSpaceOnDisks(server.Name);
                }
                else
                {
                    partitions = ServersController.GetRemainingSpaceOnMappedPartitions(server);
                }
                partitions = GetPartitionsInfos(partitions, server_disks);
                List<bool> criticals = new List<bool>();
                foreach (ServersController.VirtualizedPartition partition in partitions)
                {
                    criticals.Add(partition.Critical);
                    serverreport.Details += "Disk=" + partition.Name + "|Free=" +
                            partition.AvailableSpace;
                    if (partition.Critical)
                    {
                        serverreport.Details += "|KO";
                    }
                    else
                    {
                        serverreport.Details += "|OK";
                    }
                    serverreport.Details += "; ";
                }
                if (serverreport.Details.Length > 2)
                {
                    serverreport.Details = serverreport.Details.Substring(0, serverreport.Details.Length - 2);
                }
                serverreport.State = (!criticals.Contains(false)) ? "OK" :
                    (!criticals.Contains(true)) ? "KO" : "H-OK";
                if (ModelState.IsValid)
                {
                    db.SpaceServerReports.Add(serverreport);
                    db.SaveChanges();
                }
            }
            report.Duration = DateTime.Now.Subtract(report.DateTime);
            report.TotalChecked = report.SpaceServer_Reports.Count;
            report.TotalErrors = report.SpaceServer_Reports.Where(ser => ser.State != "OK").Count();
            if (ModelState.IsValid)
            {
                db.Entry(report).State = System.Data.Entity.EntityState.Modified;
                db.SaveChanges();
            }
            report = RegisterResults(report);
            if (ModelState.IsValid)
            {
                db.Entry(report).State = System.Data.Entity.EntityState.Modified;
                db.SaveChanges();
            }
            BuildEmail(emailId);
            Emails_Controller.AutoSend(email.Id);
            schedule.State = (schedule.Multiplicity != "Une fois") ? "Planifié" : "Terminé";
            schedule.NextExecution = Schedules_Controller.GetNextExecution(schedule);
            schedule.Executed++;
            if (ModelState.IsValid)
            {
                db.Entry(schedule).State = System.Data.Entity.EntityState.Modified;
                db.SaveChanges();
            }
            results["response"] = "OK";
            results["email"] = emailId.ToString();
            results["errors"] = "Fin d'exécution. \n" + "Erreurs d'exécution : " + results["errors"];
            Specific_Logging(new Exception("...."), "ExecuteSchedule", 3);
            return Json(results, JsonRequestBehavior.AllowGet);
        }

        private void Specific_Logging(Exception exception, string action, int level = 0)
        {
            string author = "UNKNOWN";
            if (User != null && User.Identity != null && User.Identity.Name != " ")
            {
                author = User.Identity.Name;
            }
            McoUtilities.Specific_Logging(exception, action, HomeController.SPACE_MODULE, level, author);
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
