using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Net.NetworkInformation;
using System.Runtime.InteropServices;
using System.Security.Principal;
using System.ServiceProcess;
using System.Web;
using System.Web.Mvc;
using DotNet.Highcharts;
using DotNet.Highcharts.Enums;
using DotNet.Highcharts.Helpers;
using DotNet.Highcharts.Options;
using Microsoft.Win32.TaskScheduler;
using Excel = Microsoft.Office.Interop.Excel;

namespace McoEasyTool.Controllers
{
    [AllowAnonymous]
    public class McoBesrController : Controller
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
            ViewBag.BESR_DESC_0 = McoUtilities.GetModuleDescription(HomeController.BESR_MODULE, 0);
            ViewBag.BESR_DESC_1 = McoUtilities.GetModuleDescription(HomeController.BESR_MODULE, 1);
            ViewBag.BESR_DESC_2 = McoUtilities.GetModuleDescription(HomeController.BESR_MODULE, 2);
            ViewBag.BESR_DESC_3 = McoUtilities.GetModuleDescription(HomeController.BESR_MODULE, 3);
            ViewBag.BESR_DESC_4 = McoUtilities.GetModuleDescription(HomeController.BESR_MODULE, 4);
            return View();
        }

        public ActionResult DisplaySchedules()
        {
            return View(db.BackupSchedules.OrderBy(name => name.TaskName).ToList());
        }

        public ActionResult DisplayReports()
        {
            return View(db.BackupReports.OrderByDescending(report => report.Id).ToList());
        }

        public ActionResult DisplayScheduleReports(int id)
        {
            BackupSchedule schedule = db.BackupSchedules.Find(id);
            object[] boundaries =
                McoUtilities.GetIdValues<BackupSchedule>(schedule, HomeController.OBJECT_ATTR_ID);
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
            ICollection<BackupReport> reports = db.BackupReports.Where(report => report.ScheduleId == id).ToList();
            return View(reports.OrderByDescending(report => report.Id).ToList());
        }

        public ActionResult DisplayFailedServers()
        {
            CalendarWeekRule weekRule = CalendarWeekRule.FirstDay;
            DayOfWeek firstWeekDay = DayOfWeek.Monday;
            Calendar calendar = System.Threading.Thread.CurrentThread.CurrentCulture.Calendar;
            int currentWeek = calendar.GetWeekOfYear(DateTime.Now, weekRule, firstWeekDay);

            List<BackupServer_Report> backupserverreports = null;
            try
            {
                backupserverreports = db.BackupServerReports
                    .Where(state => state.State != "OK").Where(week => week.BackupReport.WeekNumber == currentWeek)
                    .Distinct().ToList();
            }
            catch { }

            if (backupserverreports == null)
            {
                return View();
            }
            string list = "";
            foreach (BackupServer_Report serverreport in backupserverreports)
            {
                list += serverreport.BackupServer.Id + ",";
            }
            if (list.Length > 0)
            {
                list = list.Substring(0, list.Length - 1);
                BackupFailedServersUpdater(list);
            }

            try
            {
                backupserverreports = db.BackupServerReports
                    .Where(state => state.State != "OK").Where(week => week.BackupReport.WeekNumber == currentWeek).ToList();
            }
            catch { }
            if (backupserverreports == null)
            {
                return View();
            }

            return View(backupserverreports.OrderByDescending(report => report.BackupReport.LastUpdate).ToList());
        }

        public ActionResult DisplayRecipients()
        {
            return View(db.Recipients.Where(rec => rec.Module == HomeController.BESR_MODULE).ToList());
        }

        public ActionResult DisplayReportDetails(int id)
        {
            BackupReport report = db.BackupReports.Find(id);
            if (report == null)
            {
                return HttpNotFound();
            }
            object[] boundaries =
                McoUtilities.GetIdValues<BackupReport>(report, HomeController.OBJECT_ATTR_ID, true);
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
            return View(db.BackupServerReports.Where(rep => rep.BackupReportId == id).ToList());
        }

        public ActionResult DisplayReportStatistics()
        {
            BackupReport[] reports = db.BackupReports.ToArray();
            String[] reportsDate = new String[reports.Length];
            Object[] totalerrors = new Object[reports.Length];
            Object[] totalchecked = new Object[reports.Length];
            int index = 0;
            foreach (BackupReport report in reports)
            {
                reportsDate[index] = report.DateTime.ToString();
                totalerrors[index] = report.TotalErrors;
                totalchecked[index] = report.TotalChecked;
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
                    Text = "Serveurs OK",
                    X = -20
                })
                .SetSubtitle(new Subtitle
                {
                    Text = "Le nombre de serveurs OK détectés par les checks de ces derniers jours",
                    X = -20
                })
                .SetXAxis(new XAxis
                {
                    Categories = reportsDate
                })
                .SetYAxis(new YAxis
                {
                    Title = new YAxisTitle { Text = "Nombre de serveurs" },
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
                                    this.x +': '+ this.y +' serveurs';
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
                        new Series { Name = "Serveurs vérifiés", Data = new Data(totalchecked),Color = ColorTranslator.FromHtml("#12ef21") },
                        new Series { Name = "Serveurs en erreur", Data = new Data(totalerrors),Color = ColorTranslator.FromHtml("#fe5114") },
                    }
                );

            return View(chart);
        }

        public ActionResult DisplayScheduleReportDetails(int id)
        {
            BackupReport report = db.BackupReports.Find(id);
            BackupSchedule schedule = db.BackupSchedules.Find(report.ScheduleId);
            object[] boundaries =
                McoUtilities.GetIdValues<BackupReport>(report, HomeController.OBJECT_ATTR_ID, true);
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
            return View(db.BackupServerReports.Where(rep => rep.BackupReportId == id).ToList());
        }

        public ActionResult DisplayBackupServerStatistics()
        {
            BackupServer[] backupservers = db.BackupServers.ToArray();
            String[] servernames = new String[backupservers.Length];
            Object[] errorserver = new Object[backupservers.Length];
            int index = 0;
            foreach (BackupServer backupserver in backupservers)
            {
                servernames[index] = backupserver.Name;
                errorserver[index] = backupserver.BackupServer_Reports.Where(serv => serv.State != "OK").Count();
                index++;
            }
            string start = db.Reports.Min(date => date.DateTime).ToString();
            string stop = db.Reports.Max(date => date.DateTime).ToString();
            Highcharts chart = new Highcharts("chart")
                .InitChart(new Chart
                {
                    DefaultSeriesType = ChartTypes.Column,
                    MarginRight = 130,
                    MarginBottom = 25,
                    BackgroundColor = new BackColorOrGradient(ColorTranslator.FromHtml("#ded9d4")),
                    ClassName = "BackupServer"
                })
                .SetTitle(new Title
                {
                    Text = "Statistiques des défaillances",
                    X = -20
                })
                .SetSubtitle(new Subtitle
                {
                    Text = "Le nombre de rapport ayant référencé ces serveurs entre le " + start + " et le " + stop,
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

        public ActionResult BackupServerChecker(int id)
        {
            Dictionary<string, string> results = new Dictionary<string, string>();
            results["response"] = "";
            results["status"] = "";
            results["errors"] = "";

            string ExecutionErrors = "";

            BackupServer server = db.BackupServers.Find(id);
            if (server == null)
            {
                results["response"] = "KO";
                results["status"] = null;
                results["errors"] = "Fin d'exécution. \n" + "Erreurs d'exécution : Ce serveur n'existe pas ou a été supprimé.";
                return Json(results, JsonRequestBehavior.AllowGet);
            }

            try
            {
                MyApplication = new Excel.Application();
                MyApplication.Visible = false;
                MyApplication.DisplayAlerts = false;
                bool foundedFile = false;
                string[] report_files = System.IO.Directory.GetFiles(HomeController.BESR_RESULTS_FOLDER,
                    "Controle_hebdomadaire_sauvegarde_BESR_PCI_" + DateTime.Now.ToString("yyyy") + "*.xlsx", System.IO.SearchOption.TopDirectoryOnly);
                string filename = "Controle_hebdomadaire_sauvegarde_BESR_PCI_" + DateTime.Now.ToString("yyyy") + ".xlsx";
                if (report_files.Length > 0)
                {
                    filename = Directory.GetFiles(HomeController.BESR_RESULTS_FOLDER,
                        "Controle_hebdomadaire_sauvegarde_BESR_PCI_" + DateTime.Now.ToString("yyyy") + "*.xlsx", System.IO.SearchOption.TopDirectoryOnly)
                            .Select(x => new FileInfo(x))
                            .OrderByDescending(x => x.LastWriteTime).FirstOrDefault().FullName;

                    MyWorkbook = MyApplication.Workbooks.Open(filename);
                    foundedFile = true;
                }
                else
                {
                    ExecutionErrors += HomeController.BESR_RESULTS_FOLDER + "Controle_hebdomadaire_sauvegarde_BESR_PCI_" + DateTime.Now.ToString("yyyy") + "-(id).xlsx n'a pas été trouvé.\r\n";
                    MyWorkbook = MyApplication.Workbooks.Add(Excel.XlWBATemplate.xlWBATWorksheet);
                    foundedFile = false;
                }

                CalendarWeekRule weekRule = CalendarWeekRule.FirstDay;
                DayOfWeek firstWeekDay = DayOfWeek.Monday;
                Calendar calendar = System.Threading.Thread.CurrentThread.CurrentCulture.Calendar;
                int currentWeek = calendar.GetWeekOfYear(DateTime.Now, weekRule, firstWeekDay);

                string sheetName = "Semaine " + currentWeek.ToString();
                bool foundedSheet = false;
                foreach (Excel.Worksheet sheet in MyWorkbook.Sheets)
                {
                    if (sheet.Name.StartsWith(sheetName))
                    {
                        sheetName = sheet.Name;
                        foundedSheet = true;
                        break;
                    }
                }
                if (foundedSheet)
                {
                    MySheet = (Excel.Worksheet)MyWorkbook.Sheets[sheetName];
                }
                else
                {
                    MySheet = (Excel.Worksheet)MyWorkbook.Sheets.Add();
                    MySheet.Name = sheetName + " " + DateTime.Now.ToString("dd") +
                        DateTime.Now.ToString("MM") + DateTime.Now.ToString("yyyy");
                }

                MySheet.Activate();

                //Database objects init
                IQueryable<BackupReport> oldReports = db.BackupReports.Where(reportid => reportid.WeekNumber == currentWeek); //.First();
                BackupReport report;
                Email email;

                if (oldReports.Count() == 0)
                {
                    report = db.BackupReports.Create();
                    report.DateTime = DateTime.Now;
                    report.LastUpdate = DateTime.Now;
                    report.WeekNumber = currentWeek;
                    report.TotalChecked = 0;
                    report.TotalErrors = 0;
                    report.ResultPath = "";

                    email = db.Emails.Create();
                    report.Email = email;
                    email.Report = report;
                    email.Recipients = "";
                    email = Emails_Controller.SetRecipients(email, HomeController.BESR_MODULE);
                    if (ModelState.IsValid)
                    {
                        db.BackupReports.Add(report);
                        db.SaveChanges();
                        int reportNumber = db.BackupReports.Count();
                        if (reportNumber > HomeController.BESR_MAX_REPORT_NUMBER)
                        {
                            int reportNumberToDelete = reportNumber - HomeController.BESR_MAX_REPORT_NUMBER;
                            BackupReport[] reportsToDelete =
                                db.BackupReports.OrderBy(idReport => idReport.Id).Take(reportNumberToDelete).ToArray();
                            foreach (BackupReport toDeleteReport in reportsToDelete)
                            {
                                DeleteBackupReport(toDeleteReport.Id);
                            }
                        }
                    }
                    else
                    {
                        results["response"] = "KO";
                        results["status"] = null;
                        results["errors"] = "Impossible de créer un rapport dans la base de données.";
                        return Json(results, JsonRequestBehavior.AllowGet);
                    }
                }
                else
                {
                    report = oldReports.First();
                    report.DateTime = DateTime.Now;
                    email = report.Email;
                    email.Recipients = "";
                    email = Emails_Controller.SetRecipients(email, HomeController.BESR_MODULE);
                    if (ModelState.IsValid)
                    {
                        db.Entry(report).State = System.Data.Entity.EntityState.Modified;
                        db.SaveChanges();
                    }
                    else
                    {
                        results["response"] = "KO";
                        results["status"] = null;
                        results["errors"] = "Impossible de créer un rapport dans la base de données.";
                        return Json(results, JsonRequestBehavior.AllowGet);
                    }
                }

                //End database objects init

                int lastRow = MySheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row;
                ICollection<BackupServer> list = db.BackupServers.OrderBy(idServer => idServer.Pool.Id).ToArray();
                if (!foundedSheet)
                {
                    foreach (BackupServer backupserver in list)
                    {
                        Color rangeColor = ContrastColor(server.Pool.CellColor);
                        Excel.Range ActualRange = MySheet.get_Range("A" + lastRow,
                                "M" + lastRow);
                        ActualRange.Interior.Color = System.Drawing.ColorTranslator.FromHtml(backupserver.Pool.CellColor);
                        ActualRange.Font.Color = System.Drawing.ColorTranslator.FromHtml(ColorTranslator.ToHtml(rangeColor));
                        MySheet.Cells[lastRow, 1] = backupserver.Pool.Name;
                        MySheet.Cells[lastRow, 1].EntireColumn.ColumnWidth = 15;
                        MySheet.Cells[lastRow, 2] = backupserver.Name;
                        MySheet.Cells[lastRow, 2].EntireColumn.ColumnWidth = 15;
                        MySheet.Cells[lastRow, 3] = backupserver.Disks;
                        MySheet.Cells[lastRow, 3].EntireColumn.ColumnWidth = 10;
                        MySheet.Cells[lastRow, 3].EntireRow.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                        MySheet.Cells[lastRow, 4].EntireColumn.ColumnWidth = 5;
                        lastRow += 1;
                    }
                    Excel.Range styling = MySheet.get_Range("E:E", System.Type.Missing);
                    styling.EntireColumn.ColumnWidth = 60;
                }

                DateTime now = DateTime.Now;
                //DateTime yesterday = DateTime.Now.AddDays(-1);
                DateTime yesterday = DateTime.Now.AddDays(-1);
                int dayOfWeek = (int)now.Date.DayOfWeek;

                Pool pool = server.Pool;
                DayOfWeek[] week = new[] { DayOfWeek.Sunday, DayOfWeek.Monday, DayOfWeek.Tuesday,
                                     DayOfWeek.Wednesday,DayOfWeek.Thursday,DayOfWeek.Friday,
                                     DayOfWeek.Saturday};
                for (int day = -7; day < 0; day++)
                {
                    yesterday = DateTime.Now.AddDays(day);
                    if (yesterday.DayOfWeek == week[pool.BackupDay])
                    {
                        break;
                    }
                }

                //START OF SERVER CHECKING
                int row_index = 5;

                //Empty the old report on the server
                ICollection<BackupServer_Report> oldserverreports = db.BackupServerReports.Where(
                        serverreportid => serverreportid.BackupServerId == server.Id)
                        .Where(weeknum => weeknum.BackupReport.WeekNumber == currentWeek).ToList(); //.First();
                BackupServer_Report serverReport;
                if (oldserverreports == null)
                {
                    serverReport = new BackupServer_Report();
                    serverReport.BackupReport = report;
                    serverReport.BackupServer = server;
                    serverReport.Details = "";
                    serverReport.Services = "";
                    serverReport.Relaunched = "";
                }
                else
                {
                    serverReport = oldserverreports.First();
                }


                //End of Database management

                Excel.Range range = MySheet.get_Range("B1",
                    "B" + MySheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row);

                //LOOK FOR FILES
                string folderPath = HomeController.BACKUP_REMOTE_CHECK_SERVER_ROOT_FOLDER + server.Pool.Name + "\\" +
                    server.Name;
                string[] disks = server.Disks.Split(',');
                using (UNC_ACCESSOR)
                {
                    UNC_ACCESSOR.NetUseWithCredentials(folderPath,
                        HomeController.DEFAULT_USERNAME_IMPERSONNATION,
                        HomeController.DEFAULT_DOMAIN_IMPERSONNATION,
                        McoUtilities.Decrypt(HomeController.DEFAULT_PASSWORD_IMPERSONNATION));
                    try
                    {
                        string[] files = System.IO.Directory.GetFiles(folderPath, "*.*v2i");
                        if ((files == null) || (files.Count() < (disks.Count() + 1)))
                        {
                            foreach (Excel.Range cell in range)
                            {
                                if (cell.Value == server.Name)
                                {
                                    serverReport.State = "KO";
                                    serverReport.Details = "Nombre de fichiers sauvegardés incorrect. ";
                                    MySheet.Cells[cell.Row, 4] = "KO";
                                    MySheet.Cells[cell.Row, 4].Interior.Color = System.Drawing.ColorTranslator.FromHtml("#fff");
                                    MySheet.Cells[cell.Row, 4].Font.Color = System.Drawing.ColorTranslator.FromHtml("#ff0000");
                                    MySheet.Cells[cell.Row, 5] = "Nombre de fichiers insuffisant. ";
                                    break;
                                }
                            }
                            goto SERVER_NOT_OKAY;
                        }
                        else
                        {
                            if (files.Count() > (disks.Count() + 1))
                            {
                                bool testfiles = GoodFilesKepper(server, yesterday);
                                if (!testfiles)
                                {
                                    foreach (Excel.Range cell in range)
                                    {
                                        if (cell.Value == server.Name)
                                        {
                                            serverReport.State = "KO";
                                            serverReport.Details = "Nombre de fichiers sauvegardés incorrect. ";
                                            MySheet.Cells[cell.Row, 4] = "KO";
                                            MySheet.Cells[cell.Row, 4].Interior.Color = System.Drawing.ColorTranslator.FromHtml("#fff");
                                            MySheet.Cells[cell.Row, 4].Font.Color = System.Drawing.ColorTranslator.FromHtml("#ff0000");
                                            MySheet.Cells[cell.Row, 5] = "Nombre de fichiers trop important. ";
                                            break;
                                        }
                                    }
                                    goto SERVER_NOT_OKAY;
                                }
                            }
                        }
                        foreach (string disk in disks)
                        {
                            if (disk.Trim() == " " || disk.Trim() == "")
                            {
                                continue;
                            }
                            foreach (string file in files)
                            {
                                if (file.IndexOf(server.Name, StringComparison.OrdinalIgnoreCase) > 0
                                    && file.IndexOf("_" + disk.Trim() + "_", StringComparison.OrdinalIgnoreCase) > 0
                                    && file.IndexOf(".v2i", StringComparison.OrdinalIgnoreCase) > 0)
                                {
                                    DateTime lastupdate = System.IO.File.GetLastWriteTime(file);
                                    if (yesterday.Date.CompareTo(lastupdate.Date) <= 0)
                                    {
                                        foreach (Excel.Range cell in range)
                                        {
                                            if (cell.Value == server.Name)
                                            {
                                                serverReport.State = "OK";
                                                MySheet.Cells[cell.Row, 4] = "OK";
                                                MySheet.Cells[cell.Row, 4].Interior.Color = System.Drawing.ColorTranslator.FromHtml("#fff");
                                                MySheet.Cells[cell.Row, 4].Font.Color = System.Drawing.ColorTranslator.FromHtml("#00b050");
                                                MySheet.Cells[cell.Row, row_index] = file;
                                                MySheet.Cells[cell.Row, row_index].EntireColumn.ColumnWidth = 60;
                                                MySheet.Cells[cell.Row, row_index + 1] = lastupdate.ToString();
                                                MySheet.Cells[cell.Row, row_index + 1].EntireColumn.ColumnWidth = 25;
                                                row_index = row_index + 2;
                                                break;
                                            }
                                        }
                                    }
                                    else
                                    {
                                        foreach (Excel.Range cell in range)
                                        {
                                            if (cell.Value == server.Name)
                                            {
                                                serverReport.State = "KO";
                                                serverReport.Details = "Date de sauvegarde non valide";
                                                MySheet.Cells[cell.Row, 4] = "KO";
                                                MySheet.Cells[cell.Row, 4].Interior.Color = System.Drawing.ColorTranslator.FromHtml("#fff");
                                                MySheet.Cells[cell.Row, 4].Font.Color = System.Drawing.ColorTranslator.FromHtml("#ff0000");
                                                MySheet.Cells[cell.Row, 5] = "Partition " + disk + "dépassée: ";
                                                MySheet.Cells[cell.Row, row_index] = file;
                                                MySheet.Cells[cell.Row, row_index].EntireColumn.ColumnWidth = 60;
                                                MySheet.Cells[cell.Row, row_index + 1] = lastupdate.ToString();
                                                MySheet.Cells[cell.Row, row_index + 1].EntireColumn.ColumnWidth = 25;
                                                row_index = row_index + 2;
                                                break;
                                            }
                                        }
                                        goto SERVER_NOT_OKAY;
                                    }
                                    break;
                                }
                            }
                        }
                        string[] indexFiles = System.IO.Directory.GetFiles(folderPath, "*.sv2i");
                        if (indexFiles == null || indexFiles.Count() != 1)
                        {
                            foreach (Excel.Range cell in range)
                            {
                                if (cell.Value == server.Name)
                                {
                                    serverReport.State = "KO";
                                    serverReport.Details = "Fichier index manquant. ";
                                    MySheet.Cells[cell.Row, 4] = "KO";
                                    MySheet.Cells[cell.Row, 4].Interior.Color = System.Drawing.ColorTranslator.FromHtml("#fff");
                                    MySheet.Cells[cell.Row, 4].Font.Color = System.Drawing.ColorTranslator.FromHtml("#ff0000");
                                    MySheet.Cells[cell.Row, 5] = "Fichier index manquant. ";
                                    break;
                                }
                            }
                            goto SERVER_NOT_OKAY;
                        }
                        if (indexFiles[0].IndexOf(server.Name, StringComparison.OrdinalIgnoreCase) > 0)
                        {
                            DateTime lastupdate = System.IO.File.GetLastWriteTime(indexFiles[0]);
                            if (yesterday.Date.CompareTo(lastupdate.Date) <= 0)
                            {
                                foreach (Excel.Range cell in range)
                                {
                                    if (cell.Value == server.Name)
                                    {
                                        serverReport.State = "OK";
                                        MySheet.Cells[cell.Row, 4] = "OK";
                                        MySheet.Cells[cell.Row, 4].Interior.Color = System.Drawing.ColorTranslator.FromHtml("#fff");
                                        MySheet.Cells[cell.Row, 4].Font.Color = System.Drawing.ColorTranslator.FromHtml("#00b050");
                                        MySheet.Cells[cell.Row, row_index] = indexFiles[0];
                                        MySheet.Cells[cell.Row, row_index].EntireColumn.ColumnWidth = 60;
                                        MySheet.Cells[cell.Row, row_index + 1] = lastupdate.ToString();
                                        MySheet.Cells[cell.Row, row_index + 1].EntireColumn.ColumnWidth = 25;
                                        row_index = row_index + 2;
                                        break;
                                    }
                                }
                            }
                            else
                            {
                                foreach (Excel.Range cell in range)
                                {
                                    if (cell.Value == server.Name)
                                    {
                                        serverReport.State = "KO";
                                        serverReport.Details = "Date de fichier index non valide";
                                        MySheet.Cells[cell.Row, 4] = "KO";
                                        MySheet.Cells[cell.Row, 4].Interior.Color = System.Drawing.ColorTranslator.FromHtml("#fff");
                                        MySheet.Cells[cell.Row, 4].Font.Color = System.Drawing.ColorTranslator.FromHtml("#ff0000");
                                        MySheet.Cells[cell.Row, 5] = "Date de fichier index non valide. " + lastupdate.ToString();
                                        MySheet.Cells[cell.Row, row_index] = indexFiles[0];
                                        MySheet.Cells[cell.Row, row_index].EntireColumn.ColumnWidth = 60;
                                        MySheet.Cells[cell.Row, row_index + 1] = lastupdate.ToString();
                                        MySheet.Cells[cell.Row, row_index + 1].EntireColumn.ColumnWidth = 25;
                                        row_index = row_index + 2;
                                        break;
                                    }
                                }
                                goto SERVER_NOT_OKAY;
                            }
                        }
                        else
                        {
                            foreach (Excel.Range cell in range)
                            {
                                if (cell.Value == server.Name)
                                {
                                    serverReport.State = "KO";
                                    serverReport.Details = "Fichier index manquant. ";
                                    MySheet.Cells[cell.Row, 4] = "KO";
                                    MySheet.Cells[cell.Row, 4].Interior.Color = System.Drawing.ColorTranslator.FromHtml("#fff");
                                    MySheet.Cells[cell.Row, 4].Font.Color = System.Drawing.ColorTranslator.FromHtml("#ff0000");
                                    MySheet.Cells[cell.Row, 5] = "Fichier index manquant. ";
                                    break;
                                }
                            }
                            goto SERVER_NOT_OKAY;
                        }
                    }
                    catch (DirectoryNotFoundException)
                    {
                        foreach (Excel.Range cell in range)
                        {
                            if (cell.Value == server.Name)
                            {
                                serverReport.State = "KO";
                                serverReport.Details = "Répertoire Absent. ";
                                MySheet.Cells[cell.Row, 4] = "KO";
                                MySheet.Cells[cell.Row, 4].Interior.Color = System.Drawing.ColorTranslator.FromHtml("#fff");
                                MySheet.Cells[cell.Row, 4].Font.Color = System.Drawing.ColorTranslator.FromHtml("#ff0000");
                                MySheet.Cells[cell.Row, 5] = "Répertoire Absent. ";
                                break;
                            }
                        }
                        goto SERVER_NOT_OKAY;
                    }

                    catch (UnauthorizedAccessException)
                    {
                        foreach (Excel.Range cell in range)
                        {
                            if (cell.Value == server.Name)
                            {
                                serverReport.State = "KO";
                                serverReport.Details = "Accès refusé. ";
                                MySheet.Cells[cell.Row, 4] = "KO";
                                MySheet.Cells[cell.Row, 4].Interior.Color = System.Drawing.ColorTranslator.FromHtml("#fff");
                                MySheet.Cells[cell.Row, 4].Font.Color = System.Drawing.ColorTranslator.FromHtml("#ff0000");
                                MySheet.Cells[cell.Row, 5] = "Accès refusé. ";
                                break;
                            }
                        }
                        goto SERVER_NOT_OKAY;
                    }
                    catch (Exception)
                    {
                        foreach (Excel.Range cell in range)
                        {
                            if (cell.Value == server.Name)
                            {
                                serverReport.State = "KO";
                                serverReport.Details = "Erreur inconue. ";
                                MySheet.Cells[cell.Row, 4] = "KO";
                                MySheet.Cells[cell.Row, 4].Interior.Color = System.Drawing.ColorTranslator.FromHtml("#fff");
                                MySheet.Cells[cell.Row, 4].Font.Color = System.Drawing.ColorTranslator.FromHtml("#ff0000");
                                MySheet.Cells[cell.Row, 5] = "Erreur inconue. ";
                                break;
                            }
                        }
                        goto SERVER_NOT_OKAY;
                    }

                }
            //GOTO SERVER NOT OKAY
            SERVER_NOT_OKAY:
                {
                    serverReport.Services = "";
                    serverReport.Relaunched = "";
                    if (serverReport.State == "KO")
                    {
                        if (serverReport.Services == "")
                        {
                            serverReport.Services = "Services non lancés, mode Check seulement";
                        }
                        if (serverReport.Relaunched == "")
                        {
                            serverReport.Relaunched = "Non relancées, mode Check seulement";
                        }


                        try
                        {
                            Ping ping = new Ping();
                            PingOptions options = new PingOptions(64, true);
                            PingReply pingreply = ping.Send(server.Name);
                            serverReport.Ping = "Ping " + pingreply.Status.ToString();
                        }
                        catch
                        {
                            serverReport.Ping = "Ping KO";
                            if (serverReport.Services == "")
                            {
                                serverReport.Services = "Services: Ping KO";
                            }
                            if (serverReport.Relaunched == "")
                            {
                                serverReport.Relaunched = "Non relancées: Ping KO";
                            }
                        }
                        foreach (Excel.Range cell in range)
                        {
                            if (cell.Value == server.Name)
                            {
                                MySheet.Cells[cell.Row, 6] = serverReport.Ping;
                                //MySheet.Cells[cell.Row, 7] = serverReport.Services;
                                MySheet.Cells[cell.Row, 7] = serverReport.Relaunched;
                                break;
                            }
                        }
                    }
                    else
                    {
                        serverReport.Ping = "Ping OK";
                        serverReport.Services = "";
                        serverReport.Relaunched = "";
                    }
                    if (ModelState.IsValid)
                    {
                        db.BackupServerReports.Add(serverReport);
                        db.SaveChanges();
                        results["status"] = serverReport.State;
                    }
                    else
                    {
                        results["response"] = "KO";
                        results["status"] = null;
                        results["errors"] = "Erreur lors de l'enregistrement dans la base de données.";
                        Specific_Logging(new Exception("...."), "BackupServerChecker " + pool.Name, 2);
                        return Json(results, JsonRequestBehavior.AllowGet);
                    }
                }
                //END OF SERVER NOT OKAY

                //END OF SERVER CHECKING


                try
                {
                    MyWorkbook.SaveAs(HomeController.BESR_RESULTS_FOLDER + "Controle_hebdomadaire_sauvegarde_BESR_PCI_" + DateTime.Now.ToString("yyyy") + "-" +
                        report.Id.ToString() + ".xlsx",
                        Excel.XlFileFormat.xlOpenXMLWorkbook, null, null, null, null,
                        Excel.XlSaveAsAccessMode.xlNoChange, null, null, null, null, null);

                    report.ResultPath = HomeController.BESR_RESULTS_FOLDER + "Controle_hebdomadaire_sauvegarde_BESR_PCI_" + DateTime.Now.ToString("yyyy") + "-" +
                        report.Id.ToString() + ".xlsx";
                }
                catch (Exception saveException)
                {
                    ExecutionErrors += "Erreur de sauvegarde: " + saveException.Message + "\r\n";
                    Specific_Logging(saveException, "BackupServerChecker");
                    try
                    {
                        MyWorkbook.SaveAs(HomeController.BESR_RESULTS_FOLDER + "Controle_hebdomadaire_sauvegarde_BESR_PCI_" + DateTime.Now.ToString("yyyy") + ".xlsx",
                            Excel.XlFileFormat.xlOpenXMLWorkbook, null, null, null, null,
                            Excel.XlSaveAsAccessMode.xlNoChange, null, null, null, null, null);
                        report.ResultPath = HomeController.BESR_RESULTS_FOLDER + "Controle_hebdomadaire_sauvegarde_BESR_PCI_" + DateTime.Now.ToString("yyyy") + ".xlsx";
                    }
                    catch { }
                }
                report.Duration = DateTime.Now.Subtract(report.DateTime);
                report.LastUpdate = DateTime.Now;
                report.TotalChecked = report.BackupServer_Reports.Count;
                report.TotalErrors = report.BackupServer_Reports.Where(serverreport => serverreport.State != "OK").Count();
                if (ModelState.IsValid)
                {
                    db.Entry(report).State = System.Data.Entity.EntityState.Modified;
                    db.SaveChanges();
                }
                else
                {
                    results["response"] = "KO";
                    results["status"] = null;
                    results["errors"] = "Echec lors de l'enregistrement dans la base de données.";
                    return Json(results, JsonRequestBehavior.AllowGet);
                }

            }
            catch (Exception running)
            {
                try
                {
                    Specific_Logging(running, "BackupServerChecker");
                }
                catch { }
            }
            finally
            {

                McoUtilities.CloseExcel(MyApplication, MyWorkbook, MySheet);

            }
            results["response"] = "OK";
            results["errors"] = "Fin d'exécution. \n" + "Erreurs d'exécution : " + ExecutionErrors;
            return Json(results, JsonRequestBehavior.AllowGet);
        }

        public ActionResult BackupFailedServersUpdater(string list)
        {
            Dictionary<string, string> results = new Dictionary<string, string>();
            results["response"] = "";
            results["status"] = "";
            results["errors"] = "";

            string ExecutionErrors = "";
            if (list.Trim() == "")
            {
                results["response"] = "KO";
                results["status"] = null;
                results["errors"] = "Fin d'exécution. \n" + "Erreurs d'exécution : Aucun serveur sélectionné.";
                return Json(results, JsonRequestBehavior.AllowGet);
            }
            else
            {
                try
                {
                    MyApplication = new Excel.Application();
                    MyApplication.Visible = false;
                    MyApplication.DisplayAlerts = false;
                    bool foundedFile = false;
                    string[] report_files = System.IO.Directory.GetFiles(HomeController.BESR_RESULTS_FOLDER,
                        "Controle_hebdomadaire_sauvegarde_BESR_PCI_" + DateTime.Now.ToString("yyyy") + "*.xlsx", System.IO.SearchOption.TopDirectoryOnly);
                    string filename = "Controle_hebdomadaire_sauvegarde_BESR_PCI_" + DateTime.Now.ToString("yyyy") + ".xlsx";
                    if (report_files.Length > 0)
                    {
                        filename = Directory.GetFiles(HomeController.BESR_RESULTS_FOLDER,
                            "Controle_hebdomadaire_sauvegarde_BESR_PCI_" + DateTime.Now.ToString("yyyy") + "*.xlsx", System.IO.SearchOption.TopDirectoryOnly)
                                .Select(x => new FileInfo(x))
                                .OrderByDescending(x => x.LastWriteTime).FirstOrDefault().FullName;

                        MyWorkbook = MyApplication.Workbooks.Open(filename);
                        foundedFile = true;
                    }
                    else
                    {
                        ExecutionErrors += HomeController.BESR_RESULTS_FOLDER + "Controle_hebdomadaire_sauvegarde_BESR_PCI_" + DateTime.Now.ToString("yyyy") + "-(id).xlsx n'a pas été trouvé.\r\n";
                        MyWorkbook = MyApplication.Workbooks.Add(Excel.XlWBATemplate.xlWBATWorksheet);
                        foundedFile = false;
                    }

                    CalendarWeekRule weekRule = CalendarWeekRule.FirstDay;
                    DayOfWeek firstWeekDay = DayOfWeek.Monday;
                    Calendar calendar = System.Threading.Thread.CurrentThread.CurrentCulture.Calendar;
                    int currentWeek = calendar.GetWeekOfYear(DateTime.Now, weekRule, firstWeekDay);

                    string sheetName = "Semaine " + currentWeek.ToString();
                    bool foundedSheet = false;
                    foreach (Excel.Worksheet sheet in MyWorkbook.Sheets)
                    {
                        if (sheet.Name.StartsWith(sheetName))
                        {
                            sheetName = sheet.Name;
                            foundedSheet = true;
                            break;
                        }
                    }
                    if (foundedSheet)
                    {
                        MySheet = (Excel.Worksheet)MyWorkbook.Sheets[sheetName];
                    }
                    else
                    {
                        MySheet = (Excel.Worksheet)MyWorkbook.Sheets.Add();
                        MySheet.Name = sheetName + " " + DateTime.Now.ToString("dd") +
                            DateTime.Now.ToString("MM") + DateTime.Now.ToString("yyyy");
                    }

                    MySheet.Activate();

                    //Database objects init
                    IQueryable<BackupReport> oldReports = db.BackupReports.Where(reportid => reportid.WeekNumber == currentWeek); //.First();
                    BackupReport report;
                    Email email;

                    if (oldReports.Count() == 0)
                    {
                        report = db.BackupReports.Create();
                        report.DateTime = DateTime.Now;
                        report.LastUpdate = DateTime.Now;
                        report.WeekNumber = currentWeek;
                        report.TotalChecked = 0;
                        report.TotalErrors = 0;
                        report.ResultPath = "";

                        email = db.Emails.Create();
                        report.Email = email;
                        email.Report = report;
                        email.Recipients = "";
                        email = Emails_Controller.SetRecipients(email, HomeController.BESR_MODULE);
                        if (ModelState.IsValid)
                        {
                            db.BackupReports.Add(report);
                            db.SaveChanges();
                            int reportNumber = db.BackupReports.Count();
                            if (reportNumber > HomeController.BESR_MAX_REPORT_NUMBER)
                            {
                                int reportNumberToDelete = reportNumber - HomeController.BESR_MAX_REPORT_NUMBER;
                                BackupReport[] reportsToDelete =
                                    db.BackupReports.OrderBy(idReport => idReport.Id).Take(reportNumberToDelete).ToArray();
                                foreach (BackupReport toDeleteReport in reportsToDelete)
                                {
                                    DeleteBackupReport(toDeleteReport.Id);
                                }
                            }
                        }
                        else
                        {
                            results["response"] = "KO";
                            results["status"] = null;
                            results["errors"] = "Impossible de créer un rapport dans la base de données.";
                            return Json(results, JsonRequestBehavior.AllowGet);
                        }
                    }
                    else
                    {
                        report = oldReports.First();
                        report.DateTime = DateTime.Now;
                        email = report.Email;
                        email.Recipients = "";
                        email = Emails_Controller.SetRecipients(email, HomeController.BESR_MODULE);
                        if (ModelState.IsValid)
                        {
                            db.Entry(report).State = System.Data.Entity.EntityState.Modified;
                            db.SaveChanges();
                        }
                        else
                        {
                            results["response"] = "KO";
                            results["status"] = null;
                            results["errors"] = "Impossible de créer un rapport dans la base de données.";
                            return Json(results, JsonRequestBehavior.AllowGet);
                        }
                    }

                    //End database objects init

                    int lastRow = MySheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row;
                    string[] servers = list.Split(',');

                    //START OF SERVER CHECKING
                    foreach (string backupserver in servers)
                    {
                        if (backupserver.Trim() == "")
                        {
                            continue;
                        }
                        int serverId = 0;
                        if (Int32.TryParse(backupserver, out serverId))
                        {
                            BackupServer server = db.BackupServers.Find(serverId);
                            if (server == null)
                            {
                                results["errors"] += "\nLe serveur d'id " + serverId + " n'a pas été retrouvé dans la base de données.";
                                continue;
                            }
                            if (!foundedSheet)
                            {
                                Color rangeColor = ContrastColor(server.Pool.CellColor);
                                Excel.Range ActualRange = MySheet.get_Range("A" + lastRow,
                                        "M" + lastRow);
                                ActualRange.Interior.Color = System.Drawing.ColorTranslator.FromHtml(server.Pool.CellColor);
                                ActualRange.Font.Color = System.Drawing.ColorTranslator.FromHtml(ColorTranslator.ToHtml(rangeColor));
                                MySheet.Cells[lastRow, 1] = server.Pool.Name;
                                MySheet.Cells[lastRow, 1].EntireColumn.ColumnWidth = 15;
                                MySheet.Cells[lastRow, 2] = server.Name;
                                MySheet.Cells[lastRow, 2].EntireColumn.ColumnWidth = 15;
                                MySheet.Cells[lastRow, 3] = server.Disks;
                                MySheet.Cells[lastRow, 3].EntireColumn.ColumnWidth = 10;
                                MySheet.Cells[lastRow, 3].EntireRow.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                                MySheet.Cells[lastRow, 4].EntireColumn.ColumnWidth = 5;
                                lastRow += 1;
                            }

                            DateTime now = DateTime.Now;
                            //DateTime yesterday = DateTime.Now.AddDays(-1);
                            DateTime yesterday = DateTime.Now.AddDays(-1);
                            int dayOfWeek = (int)now.Date.DayOfWeek;

                            Pool pool = server.Pool;
                            DayOfWeek[] week = new[] { DayOfWeek.Sunday, DayOfWeek.Monday, DayOfWeek.Tuesday,
                                     DayOfWeek.Wednesday,DayOfWeek.Thursday,DayOfWeek.Friday,
                                     DayOfWeek.Saturday};
                            for (int day = -7; day < 0; day++)
                            {
                                yesterday = DateTime.Now.AddDays(day);
                                if (yesterday.DayOfWeek == week[pool.BackupDay])
                                {
                                    break;
                                }
                            }

                            //START OF SERVER CHECKING
                            int row_index = 5;

                            //Empty the old report on the server
                            bool newserverreport = false;
                            ICollection<BackupServer_Report> oldserverreports = db.BackupServerReports.Where(
                                    serverreportid => serverreportid.BackupServerId == server.Id)
                                    .Where(weeknum => weeknum.BackupReport.WeekNumber == currentWeek).ToList(); //.First();
                            BackupServer_Report serverReport;
                            if (oldserverreports == null)
                            {
                                newserverreport = true;
                                serverReport = new BackupServer_Report();
                                serverReport.BackupReport = report;
                                serverReport.BackupServer = server;
                                serverReport.Details = "";
                                serverReport.Services = "";
                                serverReport.Relaunched = "";
                            }
                            else
                            {
                                newserverreport = false;
                                serverReport = oldserverreports.First();
                            }


                            //End of Database management

                            Excel.Range range = MySheet.get_Range("B1",
                                "B" + MySheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row);

                            //LOOK FOR FILES
                            string folderPath = HomeController.BACKUP_REMOTE_CHECK_SERVER_ROOT_FOLDER + server.Pool.Name + "\\" +
                                server.Name;
                            string[] disks = server.Disks.Split(',');
                            using (UNC_ACCESSOR)
                            {
                                UNC_ACCESSOR.NetUseWithCredentials(folderPath,
                                    HomeController.DEFAULT_USERNAME_IMPERSONNATION,
                                    HomeController.DEFAULT_DOMAIN_IMPERSONNATION,
                                    McoUtilities.Decrypt(HomeController.DEFAULT_PASSWORD_IMPERSONNATION));
                                try
                                {
                                    string[] files = System.IO.Directory.GetFiles(folderPath, "*.*v2i");
                                    if ((files == null) || (files.Count() < (disks.Count() + 1)))
                                    {
                                        foreach (Excel.Range cell in range)
                                        {
                                            if (cell.Value == server.Name)
                                            {
                                                serverReport.State = "KO";
                                                serverReport.Details = "Nombre de fichiers sauvegardés incorrect. ";
                                                MySheet.Cells[cell.Row, 4] = "KO";
                                                MySheet.Cells[cell.Row, 4].Interior.Color = System.Drawing.ColorTranslator.FromHtml("#fff");
                                                MySheet.Cells[cell.Row, 4].Font.Color = System.Drawing.ColorTranslator.FromHtml("#ff0000");
                                                MySheet.Cells[cell.Row, 5] = "Nombre de fichiers insuffisant. ";
                                                break;
                                            }
                                        }
                                        goto SERVER_NOT_OKAY;
                                    }
                                    else
                                    {
                                        if (files.Count() > (disks.Count() + 1))
                                        {
                                            bool testfiles = GoodFilesKepper(server, yesterday);
                                            if (!testfiles)
                                            {
                                                foreach (Excel.Range cell in range)
                                                {
                                                    if (cell.Value == server.Name)
                                                    {
                                                        serverReport.State = "KO";
                                                        serverReport.Details = "Nombre de fichiers sauvegardés incorrect. ";
                                                        MySheet.Cells[cell.Row, 4] = "KO";
                                                        MySheet.Cells[cell.Row, 4].Interior.Color = System.Drawing.ColorTranslator.FromHtml("#fff");
                                                        MySheet.Cells[cell.Row, 4].Font.Color = System.Drawing.ColorTranslator.FromHtml("#ff0000");
                                                        MySheet.Cells[cell.Row, 5] = "Nombre de fichiers trop important. ";
                                                        break;
                                                    }
                                                }
                                                goto SERVER_NOT_OKAY;
                                            }
                                        }
                                    }
                                    foreach (string disk in disks)
                                    {
                                        if (disk.Trim() == " " || disk.Trim() == "")
                                        {
                                            continue;
                                        }
                                        foreach (string file in files)
                                        {
                                            if (file.IndexOf(server.Name, StringComparison.OrdinalIgnoreCase) > 0
                                                && file.IndexOf("_" + disk.Trim() + "_", StringComparison.OrdinalIgnoreCase) > 0
                                                && file.IndexOf(".v2i", StringComparison.OrdinalIgnoreCase) > 0)
                                            {
                                                DateTime lastupdate = System.IO.File.GetLastWriteTime(file);
                                                if (yesterday.Date.CompareTo(lastupdate.Date) <= 0)
                                                {
                                                    foreach (Excel.Range cell in range)
                                                    {
                                                        if (cell.Value == server.Name)
                                                        {
                                                            serverReport.State = "OK";
                                                            MySheet.Cells[cell.Row, 4] = "OK";
                                                            MySheet.Cells[cell.Row, 4].Interior.Color = System.Drawing.ColorTranslator.FromHtml("#fff");
                                                            MySheet.Cells[cell.Row, 4].Font.Color = System.Drawing.ColorTranslator.FromHtml("#00b050");
                                                            MySheet.Cells[cell.Row, row_index] = file;
                                                            MySheet.Cells[cell.Row, row_index].EntireColumn.ColumnWidth = 60;
                                                            MySheet.Cells[cell.Row, row_index + 1] = lastupdate.ToString();
                                                            MySheet.Cells[cell.Row, row_index + 1].EntireColumn.ColumnWidth = 25;
                                                            row_index = row_index + 2;
                                                            break;
                                                        }
                                                    }
                                                }
                                                else
                                                {
                                                    foreach (Excel.Range cell in range)
                                                    {
                                                        if (cell.Value == server.Name)
                                                        {
                                                            serverReport.State = "KO";
                                                            serverReport.Details = "Date de sauvegarde non valide";
                                                            MySheet.Cells[cell.Row, 4] = "KO";
                                                            MySheet.Cells[cell.Row, 4].Interior.Color = System.Drawing.ColorTranslator.FromHtml("#fff");
                                                            MySheet.Cells[cell.Row, 4].Font.Color = System.Drawing.ColorTranslator.FromHtml("#ff0000");
                                                            MySheet.Cells[cell.Row, 5] = "Partition " + disk + "dépassée: ";
                                                            MySheet.Cells[cell.Row, row_index] = file;
                                                            MySheet.Cells[cell.Row, row_index].EntireColumn.ColumnWidth = 60;
                                                            MySheet.Cells[cell.Row, row_index + 1] = lastupdate.ToString();
                                                            MySheet.Cells[cell.Row, row_index + 1].EntireColumn.ColumnWidth = 25;
                                                            row_index = row_index + 2;
                                                            break;
                                                        }
                                                    }
                                                    goto SERVER_NOT_OKAY;
                                                }
                                                break;
                                            }
                                        }
                                    }
                                    string[] indexFiles = System.IO.Directory.GetFiles(folderPath, "*.sv2i");
                                    if (indexFiles == null || indexFiles.Count() != 1)
                                    {
                                        foreach (Excel.Range cell in range)
                                        {
                                            if (cell.Value == server.Name)
                                            {
                                                serverReport.State = "KO";
                                                serverReport.Details = "Fichier index manquant. ";
                                                MySheet.Cells[cell.Row, 4] = "KO";
                                                MySheet.Cells[cell.Row, 4].Interior.Color = System.Drawing.ColorTranslator.FromHtml("#fff");
                                                MySheet.Cells[cell.Row, 4].Font.Color = System.Drawing.ColorTranslator.FromHtml("#ff0000");
                                                MySheet.Cells[cell.Row, 5] = "Fichier index manquant. ";
                                                break;
                                            }
                                        }
                                        goto SERVER_NOT_OKAY;
                                    }
                                    if (indexFiles[0].IndexOf(server.Name, StringComparison.OrdinalIgnoreCase) > 0)
                                    {
                                        DateTime lastupdate = System.IO.File.GetLastWriteTime(indexFiles[0]);
                                        if (yesterday.Date.CompareTo(lastupdate.Date) <= 0)
                                        {
                                            foreach (Excel.Range cell in range)
                                            {
                                                if (cell.Value == server.Name)
                                                {
                                                    serverReport.State = "OK";
                                                    MySheet.Cells[cell.Row, 4] = "OK";
                                                    MySheet.Cells[cell.Row, 4].Interior.Color = System.Drawing.ColorTranslator.FromHtml("#fff");
                                                    MySheet.Cells[cell.Row, 4].Font.Color = System.Drawing.ColorTranslator.FromHtml("#00b050");
                                                    MySheet.Cells[cell.Row, row_index] = indexFiles[0];
                                                    MySheet.Cells[cell.Row, row_index].EntireColumn.ColumnWidth = 60;
                                                    MySheet.Cells[cell.Row, row_index + 1] = lastupdate.ToString();
                                                    MySheet.Cells[cell.Row, row_index + 1].EntireColumn.ColumnWidth = 25;
                                                    row_index = row_index + 2;
                                                    break;
                                                }
                                            }
                                        }
                                        else
                                        {
                                            foreach (Excel.Range cell in range)
                                            {
                                                if (cell.Value == server.Name)
                                                {
                                                    serverReport.State = "KO";
                                                    serverReport.Details = "Date de fichier index non valide";
                                                    MySheet.Cells[cell.Row, 4] = "KO";
                                                    MySheet.Cells[cell.Row, 4].Interior.Color = System.Drawing.ColorTranslator.FromHtml("#fff");
                                                    MySheet.Cells[cell.Row, 4].Font.Color = System.Drawing.ColorTranslator.FromHtml("#ff0000");
                                                    MySheet.Cells[cell.Row, 5] = "Date de fichier index non valide. " + lastupdate.ToString();
                                                    MySheet.Cells[cell.Row, row_index] = indexFiles[0];
                                                    MySheet.Cells[cell.Row, row_index].EntireColumn.ColumnWidth = 60;
                                                    MySheet.Cells[cell.Row, row_index + 1] = lastupdate.ToString();
                                                    MySheet.Cells[cell.Row, row_index + 1].EntireColumn.ColumnWidth = 25;
                                                    row_index = row_index + 2;
                                                    break;
                                                }
                                            }
                                            goto SERVER_NOT_OKAY;
                                        }
                                    }
                                    else
                                    {
                                        foreach (Excel.Range cell in range)
                                        {
                                            if (cell.Value == server.Name)
                                            {
                                                serverReport.State = "KO";
                                                serverReport.Details = "Fichier index manquant. ";
                                                MySheet.Cells[cell.Row, 4] = "KO";
                                                MySheet.Cells[cell.Row, 4].Interior.Color = System.Drawing.ColorTranslator.FromHtml("#fff");
                                                MySheet.Cells[cell.Row, 4].Font.Color = System.Drawing.ColorTranslator.FromHtml("#ff0000");
                                                MySheet.Cells[cell.Row, 5] = "Fichier index manquant. ";
                                                break;
                                            }
                                        }
                                        goto SERVER_NOT_OKAY;
                                    }
                                }
                                catch (DirectoryNotFoundException)
                                {
                                    foreach (Excel.Range cell in range)
                                    {
                                        if (cell.Value == server.Name)
                                        {
                                            serverReport.State = "KO";
                                            serverReport.Details = "Répertoire Absent. ";
                                            MySheet.Cells[cell.Row, 4] = "KO";
                                            MySheet.Cells[cell.Row, 4].Interior.Color = System.Drawing.ColorTranslator.FromHtml("#fff");
                                            MySheet.Cells[cell.Row, 4].Font.Color = System.Drawing.ColorTranslator.FromHtml("#ff0000");
                                            MySheet.Cells[cell.Row, 5] = "Répertoire Absent. ";
                                            break;
                                        }
                                    }
                                    goto SERVER_NOT_OKAY;
                                }

                                catch (UnauthorizedAccessException)
                                {
                                    foreach (Excel.Range cell in range)
                                    {
                                        if (cell.Value == server.Name)
                                        {
                                            serverReport.State = "KO";
                                            serverReport.Details = "Accès refusé. ";
                                            MySheet.Cells[cell.Row, 4] = "KO";
                                            MySheet.Cells[cell.Row, 4].Interior.Color = System.Drawing.ColorTranslator.FromHtml("#fff");
                                            MySheet.Cells[cell.Row, 4].Font.Color = System.Drawing.ColorTranslator.FromHtml("#ff0000");
                                            MySheet.Cells[cell.Row, 5] = "Accès refusé. ";
                                            break;
                                        }
                                    }
                                    goto SERVER_NOT_OKAY;
                                }
                                catch (Exception)
                                {
                                    foreach (Excel.Range cell in range)
                                    {
                                        if (cell.Value == server.Name)
                                        {
                                            serverReport.State = "KO";
                                            serverReport.Details = "Erreur inconue. ";
                                            MySheet.Cells[cell.Row, 4] = "KO";
                                            MySheet.Cells[cell.Row, 4].Interior.Color = System.Drawing.ColorTranslator.FromHtml("#fff");
                                            MySheet.Cells[cell.Row, 4].Font.Color = System.Drawing.ColorTranslator.FromHtml("#ff0000");
                                            MySheet.Cells[cell.Row, 5] = "Erreur inconue. ";
                                            break;
                                        }
                                    }
                                    goto SERVER_NOT_OKAY;
                                }

                            }
                        //GOTO SERVER NOT OKAY
                        SERVER_NOT_OKAY:
                            {
                                serverReport.Services = "";
                                serverReport.Relaunched = "";
                                if (serverReport.State == "KO")
                                {
                                    if (serverReport.Services == "")
                                    {
                                        serverReport.Services = "Services non lancés, mode Check seulement";
                                    }
                                    if (serverReport.Relaunched == "")
                                    {
                                        serverReport.Relaunched = "Non relancées, mode Check seulement";
                                    }


                                    try
                                    {
                                        Ping ping = new Ping();
                                        PingOptions options = new PingOptions(64, true);
                                        PingReply pingreply = ping.Send(server.Name);
                                        serverReport.Ping = "Ping " + pingreply.Status.ToString();
                                    }
                                    catch
                                    {
                                        serverReport.Ping = "Ping KO";
                                        if (serverReport.Services == "")
                                        {
                                            serverReport.Services = "Services: Ping KO";
                                        }
                                        if (serverReport.Relaunched == "")
                                        {
                                            serverReport.Relaunched = "Non relancées: Ping KO";
                                        }
                                    }
                                    foreach (Excel.Range cell in range)
                                    {
                                        if (cell.Value == server.Name)
                                        {
                                            MySheet.Cells[cell.Row, 6] = serverReport.Ping;
                                            //MySheet.Cells[cell.Row, 7] = serverReport.Services;
                                            MySheet.Cells[cell.Row, 7] = serverReport.Relaunched;
                                            break;
                                        }
                                    }
                                }
                                else
                                {
                                    serverReport.Ping = "Ping OK";
                                    serverReport.Services = "";
                                    serverReport.Relaunched = "";
                                }
                                if (ModelState.IsValid)
                                {
                                    if (newserverreport)
                                    {
                                        db.BackupServerReports.Add(serverReport);
                                    }
                                    else
                                    {
                                        db.Entry(serverReport).State = System.Data.Entity.EntityState.Modified;
                                    }
                                    db.SaveChanges();
                                    results["status"] = serverReport.State;
                                }
                                else
                                {
                                    results["response"] = "KO";
                                    results["status"] = null;
                                    results["errors"] = "Erreur lors de l'enregistrement dans la base de données.";
                                    Specific_Logging(new Exception("...."), "CheckPool " + pool.Name, 2);
                                    return Json(results, JsonRequestBehavior.AllowGet);
                                }
                            }
                            //END OF SERVER NOT OKAY

                        }
                    }
                    //END OF SERVER CHECKING
                    Excel.Range styling = MySheet.get_Range("E:E", System.Type.Missing);
                    styling.EntireColumn.ColumnWidth = 60;

                    try
                    {
                        MyWorkbook.SaveAs(HomeController.BESR_RESULTS_FOLDER + "Controle_hebdomadaire_sauvegarde_BESR_PCI_" + DateTime.Now.ToString("yyyy") + "-" +
                            report.Id.ToString() + ".xlsx",
                            Excel.XlFileFormat.xlOpenXMLWorkbook, null, null, null, null,
                            Excel.XlSaveAsAccessMode.xlNoChange, null, null, null, null, null);

                        report.ResultPath = HomeController.BESR_RESULTS_FOLDER + "Controle_hebdomadaire_sauvegarde_BESR_PCI_" + DateTime.Now.ToString("yyyy") + "-" +
                            report.Id.ToString() + ".xlsx";
                    }
                    catch (Exception saveException)
                    {
                        ExecutionErrors += "Erreur de sauvegarde: " + saveException.Message + "\r\n";
                        try
                        {
                            MyWorkbook.SaveAs(HomeController.BESR_RESULTS_FOLDER + "Controle_hebdomadaire_sauvegarde_BESR_PCI_" + DateTime.Now.ToString("yyyy") + ".xlsx",
                                Excel.XlFileFormat.xlOpenXMLWorkbook, null, null, null, null,
                                Excel.XlSaveAsAccessMode.xlNoChange, null, null, null, null, null);
                            report.ResultPath = HomeController.BESR_RESULTS_FOLDER + "Controle_hebdomadaire_sauvegarde_BESR_PCI_" + DateTime.Now.ToString("yyyy") + ".xlsx";
                        }
                        catch { }
                    }
                    report.Duration = DateTime.Now.Subtract(report.DateTime);
                    report.LastUpdate = DateTime.Now;
                    report.TotalChecked = report.BackupServer_Reports.Count;
                    report.TotalErrors = report.BackupServer_Reports.Where(serverreport => serverreport.State != "OK").Count();
                    if (ModelState.IsValid)
                    {
                        db.Entry(report).State = System.Data.Entity.EntityState.Modified;
                        db.SaveChanges();
                        string buildOk = BuildBackupEmail(email.Id);
                        if (buildOk != "BuildOK")
                        {
                            ExecutionErrors += "Erreur lors de la mise à jour du mail \n <br />";
                        }
                    }
                    else
                    {
                        results["response"] = "KO";
                        results["status"] = null;
                        results["errors"] = "Echec lors de l'enregistrement dans la base de données.";
                        return Json(results, JsonRequestBehavior.AllowGet);
                    }

                }
                catch (Exception running)
                {
                    try
                    {
                        string log = "\r\n**************************************************\r\n";
                        log += DateTime.Now.ToString() + " : " + "BackupServerChecker Scanner General Error \r\n";
                        log += running.Message + "\r\n";
                        System.IO.File.AppendAllText(HomeController.GENERAL_LOG_FILE, log);
                    }
                    catch { }
                }
                finally
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

                }
                results["response"] = "OK";
                results["errors"] = "Fin d'exécution. \n" + "Erreurs d'exécution : " + ExecutionErrors;
                return Json(results, JsonRequestBehavior.AllowGet);
            }
        }

        public ActionResult DisplayImporter()
        {
            return View();
        }

        public ActionResult DisplayPools()
        {
            return View(db.Pools.OrderBy(poo => poo.Name).ToList());
        }

        public ActionResult DisplayPoolServers(int id)
        {
            Pool pool = db.Pools.Find(id);
            object[] boundaries =
                McoUtilities.GetIdValues<Pool>(pool, HomeController.OBJECT_ATTR_NAME);
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
            ViewBag.Message = "Gestion du Pool " + pool.Name;
            ViewBag.poolId = pool.Id;
            return View(pool.BackupServers.ToList());
        }

        public ActionResult DisplayChecker()
        {
            return View(db.Pools.OrderBy(poo => poo.Name).ToList());
        }

        public ActionResult UploadInitFile()
        {
            HttpPostedFileBase file = Request.Files[0];
            if (file != null)
                file.SaveAs(HomeController.BESR_INIT_FILE);
            ViewBag.Message = Import(true);
            return DisplayPools();
        }

        public ActionResult UpdatePoolDatabase()
        {
            IntPtr userToken = IntPtr.Zero;
            bool success = McoUtilities.LogonUser(
                HomeController.DEFAULT_USERNAME_IMPERSONNATION,
                HomeController.DEFAULT_DOMAIN_IMPERSONNATION,
                McoUtilities.Decrypt(HomeController.DEFAULT_PASSWORD_IMPERSONNATION),
                (int)McoUtilities.LogonType.LOGON32_LOGON_INTERACTIVE, //2
            (int)McoUtilities.LogonProvider.LOGON32_PROVIDER_DEFAULT, //0
            out userToken);
            Pool[] pools = db.Pools.ToArray();
            string remoteMachine = HomeController.BACKUP_REMOTE_SERVER_EXEC;
            string remoteFolder = @"\\" + remoteMachine + @"\D$\backupbesr\";
            string modreport = "";
            Dictionary<int, ServersController.VirtualizedServer> FOREST = ServersController.GetInformationsFromForestDomains();
            ReftechServers[] REFTECH_SERVERS = null;
            try
            {
                REFTECH_SERVERS = db.ReftechServers.ToArray();
            }
            catch { }
            foreach (Pool pool in pools)
            {
                string initFile = "Parametres" + pool.Name + ".ini";
                using (WindowsIdentity.Impersonate(userToken))
                {
                    try
                    {
                        string[] lines = System.IO.File.ReadAllLines(remoteFolder + initFile);
                        foreach (string line in lines)
                        {
                            if (line.Trim() != "")
                            {
                                string[] infos = line.Split(';');
                                if (infos.Length >= 2 && infos[0].Trim().IndexOf("#") != 0)
                                {
                                    string servername = infos[0].Substring(infos[0].IndexOf("=") + 2).Trim().ToUpper();
                                    string partition = infos[1].Substring(infos[1].IndexOf("=") + 2, 1).Trim().ToUpper();
                                    List<BackupServer> servers = pool.BackupServers.Where(name => name.Name == servername).ToList();

                                    //SERVER EXISTS
                                    if (servers.Count == 1)
                                    {
                                        BackupServer server = servers.First();
                                        //SERVER ALREADY HAS PARTITION
                                        if (server.Disks.IndexOf(partition) != -1)
                                        {
                                            continue;
                                        }
                                        //SERVER HASN'T PARTITION 
                                        else
                                        {
                                            server.Disks += ", " + partition;
                                            if (ModelState.IsValid)
                                            {
                                                db.Entry(server).State = System.Data.Entity.EntityState.Modified;
                                                db.SaveChanges();
                                                modreport += "--------------------------------------------------------------\r\n";
                                                modreport += "Partition " + partition + " ajoutée à Pool " + pool.Name + " | Serveur " + server.Name + " : " + DateTime.Now.ToString() + "\r\n";
                                                UpdateBackupServerFile(server, "disk");
                                            }
                                            else
                                            {
                                                modreport += "--------------------------------------------------------------\r\n";
                                                modreport += "Erreur lors de l'ajout de la partition " + partition + " à Pool " + pool.Name + " | Serveur " + server.Name + " : " + DateTime.Now.ToString() + "\r\n";
                                                continue;
                                            }
                                        }
                                    }
                                    //SERVER DOESN'T EXIST
                                    else
                                    {
                                        BackupServer server = db.BackupServers.Create();
                                        server.Name = servername;
                                        ServersController.VirtualizedServer_Result virtual_server = ServersController.GetServerInformations(FOREST, REFTECH_SERVERS, server, HomeController.BESR_MODULE, false);
                                        server = virtual_server.BESR_Server;
                                        server.Disks = partition;
                                        server.Pool = pool;
                                        if (ModelState.IsValid)
                                        {
                                            db.BackupServers.Add(server);
                                            db.SaveChanges();
                                            modreport += "--------------------------------------------------------------\r\n";
                                            modreport += "Pool " + pool.Name + " | Serveur " + server.Name + " ajouté : " + DateTime.Now.ToString() + "\r\n";
                                            UpdateBackupServerFile(server, "server");
                                        }
                                        else
                                        {
                                            modreport += "--------------------------------------------------------------\r\n";
                                            modreport += "Erreur lors de l'ajout de Pool " + pool.Name + " | Serveur " + server.Name + " : " + DateTime.Now.ToString() + "\r\n";
                                            continue;
                                        }
                                    }
                                }
                                else
                                {
                                    //ANORMAL LINE
                                    continue;
                                }
                            }
                            else
                            {
                                //EMPTY LINE
                                continue;
                            }
                        }
                    }
                    catch (Exception exception)
                    {
                        Specific_Logging(exception, "UpdatePoolDatabase");
                    }
                }
                try
                {
                    string log = "\r\n**************************************************\r\n";
                    log += DateTime.Now.ToString() + " : " + "Pool UpdateDatabase : General Error \r\n";
                    log += modreport;
                    System.IO.File.AppendAllText(HomeController.BESR_AUTO_UPDATE_LOG_FILE, log);
                }
                catch { }
            }
            return View();
        }

        public ActionResult ServiceStopper(int id)
        {
            BackupServer server = db.BackupServers.Find(id);
            string resultat = "@******************************************** \n  " + server.Name + "\n ";
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
                throw new Exception("Logon user failed");
            }
            using (WindowsIdentity.Impersonate(userToken))
            {
                ServiceController[] services = ServiceController.GetServices(server.Name);
                foreach (ServiceController service in services)
                {
                    try
                    {
                        if ((service.ServiceName == "SymSnapService") ||
                            (service.ServiceName == "Backup Exec System Recovery"))
                        {
                            ServiceSpecialController targetedService = new ServiceSpecialController(service.ServiceName, server.Name);
                            if (targetedService.Status != ServiceControllerStatus.Stopped)
                            {
                                targetedService.StartupType = "Manual";
                                targetedService.Stop();
                                targetedService.WaitForStatus(ServiceControllerStatus.Stopped);
                            }
                            resultat += "----------------------------- \n ";
                            resultat += targetedService.ServiceName + ": " + targetedService.Status + ": " + targetedService.StartupType + "\n ";
                        }
                    }
                    catch (Exception exception)
                    {
                        Specific_Logging(exception, "ServiceStopper " + server.Name);
                        resultat += "----------------------------- \n ";
                        resultat += "Un problème est survenu lors de l'arrêt de " + service.ServiceName + ": \n ";
                        resultat += exception.Message + ": \n ";
                    }
                }
            }
            try
            {
                using (TaskService taskservice = new TaskService())
                {
                    TaskCollection tasks = taskservice.RootFolder.Tasks;
                    foreach (Task task in tasks)
                    {
                        if (task.Name.IndexOf(server.Name) != -1)
                        {
                            taskservice.RootFolder.DeleteTask(task.Name, false);
                            break;
                        }
                    }
                }
            }
            catch { }
            Specific_Logging(new Exception("...."), "ServiceStopper " + server.Name, 3);
            return View();
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
                    string taskname = HomeController.BESR_MODULE + " AutoCheck ";

                    BackupSchedule schedule = db.BackupSchedules.Create();
                    schedule.CreationTime = DateTime.Now;
                    schedule.NextExecution = scheduled;
                    schedule.Generator = User.Identity.Name;
                    schedule.Multiplicity = multiplicity;
                    schedule.Executed = 0;
                    schedule.State = "Planifié";
                    schedule.Module = HomeController.BESR_MODULE;
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
                            schedule.AutoRelaunch = true;
                        }
                        else
                        {
                            schedule.AutoRelaunch = false;
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
            BackupSchedule schedule = db.BackupSchedules.Find(id);
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
                    string taskname = HomeController.BESR_MODULE + " AutoCheck ";

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
                            schedule.AutoRelaunch = true;
                        }
                        else
                        {
                            schedule.AutoRelaunch = false;
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
            BackupSchedule schedule = db.BackupSchedules.Find(id);
            if (schedule.Reports.Count != 0)
            {
                List<BackupReport> reports = db.BackupReports.Where(rep => rep.ScheduleId == schedule.Id).ToList();
                foreach (Report report in reports)
                {
                    DeleteBackupReport(report.Id);
                    db.Entry(schedule).State = System.Data.Entity.EntityState.Modified;
                    db.SaveChanges();
                }
            }

            string result = Schedules_Controller.Delete(schedule);
            schedule = db.BackupSchedules.Find(id);
            if (result == "La tâche a été correctement supprimée")
            {
                db.BackupSchedules.Remove(schedule);
                db.SaveChanges();
            }
            Specific_Logging(new Exception("...."), "DeleteSchedule", 3);
            return result;
        }

        public string ReSendLastEmail(int id)
        {
            BackupSchedule schedule = db.BackupSchedules.Find(id);
            if (schedule != null)
            {
                if (db.Reports.Where(report => report.ScheduleId == id).Count() != 0)
                {
                    BackupReport report = db.BackupReports.Where(rep => rep.ScheduleId == id).OrderByDescending(rep => rep.Id).First();
                    return Reports_Controller.ReSend(report.Id);
                }
                return "Cette tâche planifiée n'a pour l'instant généré aucun rapport, ou alors ils ont été supprimés.";

            }
            Specific_Logging(new Exception("...."), "ReSendLAstEmail", 3);
            return "Cette tâche planifiée n'a pas été trouvée dans la base de données.";
        }

        public string ViewEmails(int id)
        {
            BackupReport report = db.BackupReports.Find(id);
            if (report != null)
            {
                return Reports_Controller.ViewEmail(report.Id);
            }
            return "Ce rapport n'a pas été retrouvé dans la base de données.";
        }

        public string ReSendEmail(int id)
        {
            BackupReport report = db.BackupReports.Find(id);
            if (report != null)
            {
                return Reports_Controller.ReSend(report.Id);
            }
            Specific_Logging(new Exception("...."), "ReSend", 2);
            return "Le rapport n'a pas été retrouvé dans la base de données.";
        }

        public string DownloadReport(int id)
        {
            BackupReport report = db.BackupReports.Find(id);
            if (report == null)
            {
                return HttpNotFound().ToString();
            }
            return Reports_Controller.Download(report.Id);
        }

        public string DeleteBackupReport(int id)
        {
            try
            {
                BackupReport report = db.BackupReports.Find(id);
                Email email = report.Email;

                List<BackupServer_Report> backupserverreports = db.BackupServerReports.ToList();
                foreach (BackupServer_Report backupserverreport in backupserverreports)
                {
                    if (backupserverreport.BackupReport == report)
                    {
                        db.BackupServerReports.Remove(backupserverreport);
                        db.SaveChanges();
                    }
                }
                Specific_Logging(new Exception("...."), "DeleteBackupReport", 3);
                return Reports_Controller.Delete(report.Id);
            }
            catch (Exception exception)
            {
                Specific_Logging(exception, "DeleteBackupReport");
                return "Une erreur est surveunue lors de la suppression" +
                    exception.Message;
            }
        }

        public string Purge()
        {
            string message = "";
            List<BackupReport> reports = db.BackupReports.Where(rep => rep.Duration == null || rep.ResultPath == null).ToList();
            foreach (BackupReport report in reports)
            {
                message += "Rapport " + report.DateTime + " supprimé";
                Email email = (report.Email != null) ? report.Email : null;
                List<BackupServer_Report> backupserverreports = report.BackupServer_Reports.ToList();
                foreach (BackupServer_Report backupserverreport in backupserverreports)
                {
                    db.BackupServerReports.Remove(backupserverreport);
                }
                db.SaveChanges();
                Reports_Controller.Delete(report.Id);
            }
            Specific_Logging(new Exception("...."), "Purge", 3);
            return message;
        }

        public string ServiceLauncher(int id)
        {
            BackupServer server = db.BackupServers.Find(id);
            string resultat = "@******************************************** \n  " + server.Name + "\n ";
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
                return "Logon user failed";
            }

            using (WindowsIdentity.Impersonate(userToken))
            {
                ServiceController[] services = ServiceController.GetServices(server.Name);
                foreach (ServiceController service in services)
                {
                    try
                    {
                        if ((service.ServiceName == "SymSnapService") ||
                            (service.ServiceName == "Backup Exec System Recovery"))
                        {
                            ServiceSpecialController targetedService = new ServiceSpecialController(service.ServiceName, server.Name);
                            if (targetedService.Status != ServiceControllerStatus.Running)
                            {
                                targetedService.StartupType = "Manual";
                                targetedService.Start();
                                targetedService.WaitForStatus(ServiceControllerStatus.Running);
                            }
                            resultat += "----------------------------- \n ";
                            resultat += targetedService.ServiceName + ": " + targetedService.Status + ": " + targetedService.StartupType + "\n ";
                        }
                    }
                    catch (Exception exception)
                    {
                        Specific_Logging(exception, "ServiceLauncher " + server.Name);
                        resultat += "----------------------------- \n ";
                        resultat += "Un problème est survenu lors du lancement de " + service.ServiceName + ": \n ";
                        resultat += exception.Message + ": \n ";
                    }
                }
            }
            Specific_Logging(new Exception("...."), "ServiceLauncher " + server.Name, 3);
            return resultat;
        }

        public string BackupExecLauncher(int id)
        {
            string launched = "KO";
            BackupServer server = db.BackupServers.Find(id);
            Pool pool = db.Pools.Find(server.PoolId);
            if (server == null || pool == null)
            {
                return HttpNotFound().ToString();
            }
            string result = "@******************************************** \n  " + server.Name + "\n ";
            try
            {
                System.Security.Principal.WindowsImpersonationContext impersonationContext;
                impersonationContext = ((System.Security.Principal.WindowsIdentity)User.Identity).Impersonate();
                Process process = new Process();
                process.StartInfo.UseShellExecute = false;
                process.StartInfo.LoadUserProfile = true;
                process.StartInfo.RedirectStandardOutput = true;
                process.StartInfo.RedirectStandardError = true;
                DateTime time = DateTime.Now.AddMinutes(3);
                string hours = time.ToString("HH:mm");
                string besrArguments = "";
                if (server.Version == "Windows 2000")
                {
                    besrArguments += "BackupBESR853.vbs " + server.Name +
                    " Parametres" + pool.Name + ".ini";
                }
                else
                {
                    besrArguments += "BackupBESR2010.vbs " + server.Name +
                        " Parametres" + pool.Name + ".ini";
                }
                process.StartInfo.Verb = "Runas";
                process.StartInfo.FileName = "cmd";
                process.StartInfo.Arguments = "/c powershell D:\\McoEasyTool\\BatchFiles\\BESRScheduler.ps1 " + hours + " " + besrArguments + "";
                process.Start();
                impersonationContext.Undo();
                launched = "OK ";
                process.WaitForExit();
                launched += "Exécution de la sauvegarde lancée. \n ";
                string folder_Path = GetFolderPath(server);
                IntPtr userToken = IntPtr.Zero;
                try
                {
                    bool success = McoUtilities.LogonUser(
                            HomeController.DEFAULT_USERNAME_IMPERSONNATION,
                            HomeController.DEFAULT_DOMAIN_IMPERSONNATION,
                            McoUtilities.Decrypt(HomeController.DEFAULT_PASSWORD_IMPERSONNATION),
                            (int)McoUtilities.LogonType.LOGON32_LOGON_INTERACTIVE, //2
                            (int)McoUtilities.LogonProvider.LOGON32_PROVIDER_DEFAULT, //0
                            out userToken);

                    string[] filePaths = System.IO.Directory.GetFiles(folder_Path);
                    foreach (string filePath in filePaths)
                    {
                        System.IO.File.Delete(filePath);
                    }
                }
                catch (System.IO.DirectoryNotFoundException)
                {
                    System.IO.Directory.CreateDirectory(folder_Path);
                }
                catch { }

                try
                {
                    BackupServer_Report backupserverreport = null;
                    if (db.BackupServerReports
                        .Where(serverreport => serverreport.BackupServer.Id == server.Id)
                        .OrderByDescending(serverreport => serverreport.BackupReport.Id).Count() != 0)
                    {
                        backupserverreport = db.BackupServerReports
                            .Where(serverreport => serverreport.BackupServer.Id == server.Id)
                            .OrderByDescending(serverreport => serverreport.BackupReport.Id)
                            .First();
                    }
                    if (backupserverreport != null)
                    {
                        backupserverreport.Services = "OK";
                        backupserverreport.Relaunched = "Relancées";
                        if (ModelState.IsValid)
                        {
                            db.Entry(backupserverreport).State = System.Data.Entity.EntityState.Modified;
                            db.SaveChanges();
                        }
                    }
                }
                catch (Exception exception)
                {
                    Specific_Logging(exception, "BackupExecLauncher " + server.Name);
                    result += "Exceptions survenues lors de la mise à jour de la BD: " + exception.Message + " \n";
                }
            }
            catch (Exception exception)
            {
                Specific_Logging(exception, "BackupExecLauncher " + server.Name);
                launched = "KO ";
                result += "Exceptions survenues: " + exception.Message + " \n";
            }

            string scheduled = ServiceStopperScheduler(server.Id);
            if (scheduled != "OK")
            {
                result += "Sauvegarde " + launched + ": planification arrêt services KO  \n ";
            }
            else
            {
                result += "Sauvegarde " + launched + ": Arrêt des services planifié en fin d'exécution  \n ";
            }
            Specific_Logging(new Exception("...."), "BackupExecLauncher " + server.Name, 3);
            return result;
        }

        public string PoolBackupExecLauncher(int id)
        {
            string result = "";
            Pool pool = db.Pools.Find(id);
            if (pool == null)
            {
                return HttpNotFound().ToString();
            }
            BackupServer[] servers = pool.BackupServers.ToArray();
            foreach (BackupServer server in servers)
            {
                string services = ServiceLauncher(server.Id);
                string BESROK = "Backup Exec System Recovery: Running: Manual";
                string SYMOK = "SymSnapService: Running: Manual";
                if ((services.IndexOf(BESROK, StringComparison.OrdinalIgnoreCase) > 0) &&
                    services.IndexOf(SYMOK, StringComparison.OrdinalIgnoreCase) > 0)
                {
                    result += BackupExecLauncher(server.Id) + "\n";
                }
                else
                {
                    result += "Services: " + services + "\n";
                }
            }
            Specific_Logging(new Exception("...."), "PoolBackupExecLauncher " + pool.Name, 3);
            return result;
        }

        public string BackupExecLauncherServers()
        {
            string result = "";
            string selectedServers = Request.Form["selectedServers"].ToString();
            string[] servers = selectedServers.Split(',');
            foreach (string server in servers)
            {
                if (server.Trim() == "")
                {
                    continue;
                }

                int serverId = 0;
                if (Int32.TryParse(server, out serverId))
                {
                    BackupServer backupserver = db.BackupServers.Find(serverId);
                    if (backupserver == null)
                    {
                        return HttpNotFound().ToString();
                    }
                    result += "-------------------------------------- \n";
                    result += backupserver.Name + " : ";
                    string services = ServiceLauncher(serverId);
                    string BESROK = "Backup Exec System Recovery: Running: Manual";
                    string SYMOK = "SymSnapService: Running: Manual";
                    if ((services.IndexOf(BESROK, StringComparison.OrdinalIgnoreCase) > 0) &&
                        services.IndexOf(SYMOK, StringComparison.OrdinalIgnoreCase) > 0)
                    {
                        BackupServer_Report serverreport = backupserver.BackupServer_Reports
                            .OrderByDescending(id => id.Id).First();
                        serverreport.Services = "OK";
                        string relaunched = BackupExecLauncher(serverId);
                        if (relaunched.IndexOf("Sauvegarde OK") != -1)
                        {
                            serverreport.Relaunched = "Relancées";
                            Specific_Logging(new Exception("...."), "BackupExecLauncherServers " + backupserver.Name, 3);
                        }
                        result += serverreport.Relaunched + "\n";
                        if (ModelState.IsValid)
                        {
                            db.Entry(serverreport);
                            db.SaveChanges();
                        }
                    }
                    else
                    {
                        result += "Erreur de lancement des services \n";
                    }
                }
            }
            Specific_Logging(new Exception("...."), "BackupExecLauncherServers", 3);
            return result;
        }

        public string ServiceStopperScheduler(int id)
        {
            BackupServer server = db.BackupServers.Find(id);
            string success = "";
            using (TaskService taskservice = new TaskService())
            {
                try
                {
                    TaskDefinition taskdefinition = taskservice.NewTask();
                    taskdefinition.RegistrationInfo.Description = "Arrêt des services BESR et SymSnap";
                    Trigger trigger = Trigger.CreateTrigger(TaskTriggerType.Time);
                    trigger.StartBoundary = DateTime.Now.AddMinutes(45);
                    taskdefinition.Triggers.Add(trigger);

                    taskdefinition.Principal.UserId = HomeController.DEFAULT_DOMAIN_AND_USERNAME_IMPERSONNATION;
                    taskdefinition.Settings.WakeToRun = true;
                    taskdefinition.Settings.StartWhenAvailable = true;
                    taskdefinition.Settings.StopIfGoingOnBatteries = false;
                    taskdefinition.Settings.DeleteExpiredTaskAfter.Add(new TimeSpan(0, 15, 0));
                    taskdefinition.Actions.Add(new ExecAction(HomeController.MCO_SCHEDULER_TOOL, "BESR STOP_SERVICE " + server.Id.ToString()));
                    //taskdefinition.Actions.Add(new ExecAction(HomeController.AUTO_TOOL_PATH, "BESR STOP_SERVICE " + server.Id.ToString(), null));
                    //taskservice.RootFolder.RegisterTaskDefinition("StopBesrServices" + server.Name, taskdefinition);
                    taskservice.RootFolder.RegisterTaskDefinition("StopBesrServices " + server.Name, taskdefinition,
                    TaskCreation.CreateOrUpdate, HomeController.DEFAULT_DOMAIN_AND_USERNAME_IMPERSONNATION,
                    McoUtilities.Decrypt(HomeController.DEFAULT_PASSWORD_IMPERSONNATION));
                    success += "OK";
                }
                catch (Exception exception)
                {
                    Specific_Logging(exception, "ServiceStopperScheduler " + server.Name);
                    success += "KO: " + exception.Message;
                }
                Specific_Logging(new Exception("...."), "ServiceStopperScheduler " + server.Name, 3);
                return success;
            }
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

        public bool GoodFilesKepper(BackupServer server, DateTime yesterday)
        {
            string[] disks = server.Disks.Split(',');
            string folderPath = HomeController.BACKUP_REMOTE_CHECK_SERVER_ROOT_FOLDER + server.Pool.Name + "\\" + server.Name;
            Dictionary<string, DateTime> goodfiles = new Dictionary<string, DateTime>();
            try
            {
                IntPtr userToken = IntPtr.Zero;
                bool success = McoUtilities.LogonUser(
                  HomeController.DEFAULT_USERNAME_IMPERSONNATION,
                  HomeController.DEFAULT_DOMAIN_IMPERSONNATION,
                  McoUtilities.Decrypt(HomeController.DEFAULT_PASSWORD_IMPERSONNATION),
                  (int)McoUtilities.LogonType.LOGON32_LOGON_INTERACTIVE, //2
                  (int)McoUtilities.LogonProvider.LOGON32_PROVIDER_DEFAULT, //0
                  out userToken);

                string[] files = System.IO.Directory.GetFiles(folderPath, "*.*v2i");
                foreach (string disk in disks)
                {
                    if (disk.Trim() == " " || disk.Trim() == "")
                    {
                        continue;
                    }
                    foreach (string file in files)
                    {
                        if (file.IndexOf(server.Name, StringComparison.OrdinalIgnoreCase) > 0
                            && file.IndexOf("_" + disk.Trim() + "_", StringComparison.OrdinalIgnoreCase) > 0
                            && file.IndexOf(".v2i", StringComparison.OrdinalIgnoreCase) > 0)
                        {
                            DateTime lastupdate = System.IO.File.GetLastWriteTime(file);
                            if (yesterday.Date.CompareTo(lastupdate.Date) <= 0)
                            {
                                goodfiles.Add(file, lastupdate);
                            }
                        }
                    }
                }
                if (goodfiles.Count != disks.Length)
                {
                    return false;
                }
                string[] indexFiles = System.IO.Directory.GetFiles(folderPath, "*.sv2i");
                if (indexFiles == null || indexFiles.Count() != 1)
                {
                    return false;
                }
                if (indexFiles[0].IndexOf(server.Name, StringComparison.OrdinalIgnoreCase) > 0)
                {
                    DateTime lastupdate = System.IO.File.GetLastWriteTime(indexFiles[0]);
                    if (yesterday.Date.CompareTo(lastupdate.Date) <= 0)
                    {
                        goodfiles.Add(indexFiles[0], lastupdate);
                    }
                }
                Dictionary<string, DateTime> toDelete = new Dictionary<string, DateTime>();
                foreach (string file in files)
                {
                    DateTime lastupdate = System.IO.File.GetLastWriteTime(file);
                    if (goodfiles.ContainsKey(file) && goodfiles[file] == lastupdate)
                    {
                        continue;
                    }
                    else
                    {
                        toDelete.Add(file, lastupdate);
                    }
                }

                foreach (KeyValuePair<string, DateTime> badfile in toDelete)
                {
                    System.IO.File.Delete(badfile.Key);
                }
                return true;
            }
            catch (Exception exception)
            {
                Specific_Logging(exception, "GoodFilesKeeper");
            }
            return false;
        }

        public JsonResult BesrChecker()
        {
            //UpdatePoolDatabase();
            Dictionary<string, string> results = new Dictionary<string, string>();
            results["response"] = "";
            results["email"] = "";
            results["errors"] = "";
            results["pools"] = "";
            results["autobesr"] = "";

            bool autobesr = false;
            string[] poolList;
            List<Pool> SelectedPools = new List<Pool>();
            try
            {
                bool.TryParse(Request.Form["autobesr"], out autobesr);
                results["autobesr"] = autobesr.ToString();
                results["pools"] = Request.Form["list"].ToString();
                poolList = results["pools"].Split(';');
                foreach (string poolId in poolList)
                {
                    int id = 0;
                    Int32.TryParse(poolId.Split('-')[1], out id);
                    Pool pool = db.Pools.Find(id);
                    if (pool != null)
                    {
                        SelectedPools.Add(pool);
                    }
                }
            }
            catch (Exception exception)
            {
                results["response"] = "Une erreur est survenue lors de la sélection des Pools";
                results["email"] = "";
                results["errors"] = exception.Message;
                return Json(results, JsonRequestBehavior.AllowGet);
            }
            if (SelectedPools.Count == 0)
            {
                results["response"] = "Aucun Pool n'a été sélectionné dans la base de données";
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
                MyApplication.DisplayAlerts = false;
                bool foundedFile = false;
                string[] report_files = System.IO.Directory.GetFiles(HomeController.BESR_RESULTS_FOLDER,
                    "Controle_hebdomadaire_sauvegarde_BESR_PCI_" + DateTime.Now.ToString("yyyy") + "*.xlsx", System.IO.SearchOption.TopDirectoryOnly);
                string filename = "Controle_hebdomadaire_sauvegarde_BESR_PCI_" + DateTime.Now.ToString("yyyy") + ".xlsx";
                if (report_files.Length > 0)
                {
                    filename = Directory.GetFiles(HomeController.BESR_RESULTS_FOLDER,
                        "Controle_hebdomadaire_sauvegarde_BESR_PCI_" + DateTime.Now.ToString("yyyy") + "*.xlsx", System.IO.SearchOption.TopDirectoryOnly)
                            .Select(x => new FileInfo(x))
                            .OrderByDescending(x => x.LastWriteTime).FirstOrDefault().FullName;

                    MyWorkbook = MyApplication.Workbooks.Open(filename);
                    foundedFile = true;
                }
                else
                {
                    ExecutionErrors += HomeController.BESR_RESULTS_FOLDER + "Controle_hebdomadaire_sauvegarde_BESR_PCI_" + DateTime.Now.ToString("yyyy") + "-(id).xlsx n'a pas été trouvé.\r\n";
                    MyWorkbook = MyApplication.Workbooks.Add(Excel.XlWBATemplate.xlWBATWorksheet);
                    foundedFile = false;
                }

                CalendarWeekRule weekRule = CalendarWeekRule.FirstDay;
                DayOfWeek firstWeekDay = DayOfWeek.Monday;
                Calendar calendar = System.Threading.Thread.CurrentThread.CurrentCulture.Calendar;
                int currentWeek = calendar.GetWeekOfYear(DateTime.Now, weekRule, firstWeekDay);

                string sheetName = "Semaine " + currentWeek.ToString();
                bool foundedSheet = false;
                foreach (Excel.Worksheet sheet in MyWorkbook.Sheets)
                {
                    if (sheet.Name.StartsWith(sheetName))
                    {
                        sheetName = sheet.Name;
                        foundedSheet = true;
                        break;
                    }
                }
                if (foundedSheet)
                {
                    MySheet = (Excel.Worksheet)MyWorkbook.Sheets[sheetName];
                }
                else
                {
                    MySheet = (Excel.Worksheet)MyWorkbook.Sheets.Add();
                    MySheet.Name = sheetName + " " + DateTime.Now.ToString("dd") +
                        DateTime.Now.ToString("MM") + DateTime.Now.ToString("yyyy");
                }

                MySheet.Activate();

                //Database objects init
                IQueryable<BackupReport> oldReports = db.BackupReports.Where(reportid => reportid.WeekNumber == currentWeek); //.First();
                BackupReport report;
                Email email;

                if (oldReports.Count() == 0)
                {
                    report = db.BackupReports.Create();
                    report.DateTime = DateTime.Now;
                    report.LastUpdate = DateTime.Now;
                    report.WeekNumber = currentWeek;
                    report.TotalChecked = 0;
                    report.TotalErrors = 0;
                    report.ResultPath = "";
                    report.Author = User.Identity.Name;
                    report.Module = HomeController.BESR_MODULE;

                    email = db.Emails.Create();
                    report.Email = email;
                    email.Module = HomeController.BESR_MODULE;
                    email.Report = report;
                    email.Recipients = "";
                    email = Emails_Controller.SetRecipients(email, HomeController.BESR_MODULE);
                    if (ModelState.IsValid)
                    {
                        db.BackupReports.Add(report);
                        db.SaveChanges();
                        emailId = report.Email.Id;
                        int reportNumber = db.BackupReports.Count();
                        if (reportNumber > HomeController.BESR_MAX_REPORT_NUMBER)
                        {
                            int reportNumberToDelete = reportNumber - HomeController.BESR_MAX_REPORT_NUMBER;
                            BackupReport[] reportsToDelete =
                                db.BackupReports.OrderBy(id => id.Id).Take(reportNumberToDelete).ToArray();
                            foreach (BackupReport toDeleteReport in reportsToDelete)
                            {
                                DeleteBackupReport(toDeleteReport.Id);
                            }
                        }
                    }
                    else
                    {
                        results["response"] = "KO";
                        results["email"] = null;
                        results["errors"] = "Impossible de créer un rapport dans la base de données.";
                        goto CHECK_FINALIZER;
                    }
                }
                else
                {
                    report = oldReports.First();
                    report.DateTime = DateTime.Now;
                    email = report.Email;
                    email.Recipients = "";
                    email = Emails_Controller.SetRecipients(email, HomeController.BESR_MODULE);
                    if (ModelState.IsValid)
                    {
                        db.Entry(report).State = System.Data.Entity.EntityState.Modified;
                        db.SaveChanges();
                        emailId = report.Email.Id;
                    }
                    else
                    {
                        results["response"] = "KO";
                        results["email"] = null;
                        results["errors"] = "Impossible de créer un rapport dans la base de données.";
                        goto CHECK_FINALIZER;
                    }
                }

                //End database objects init
                int lastRow = MySheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row;
                ICollection<BackupServer> list = db.BackupServers.OrderBy(id => id.Pool.Id).ToArray();
                if (!foundedSheet)
                {
                    foreach (BackupServer server in list)
                    {
                        Color rangeColor = ContrastColor(server.Pool.CellColor);
                        Excel.Range ActualRange = MySheet.get_Range("A" + lastRow,
                                "M" + lastRow);
                        ActualRange.Interior.Color = System.Drawing.ColorTranslator.FromHtml(server.Pool.CellColor);
                        ActualRange.Font.Color = System.Drawing.ColorTranslator.FromHtml(ColorTranslator.ToHtml(rangeColor));
                        MySheet.Cells[lastRow, 1] = server.Pool.Name;
                        MySheet.Cells[lastRow, 1].EntireColumn.ColumnWidth = 15;
                        MySheet.Cells[lastRow, 2] = server.Name;
                        MySheet.Cells[lastRow, 2].EntireColumn.ColumnWidth = 15;
                        MySheet.Cells[lastRow, 3] = server.Disks;
                        MySheet.Cells[lastRow, 3].EntireColumn.ColumnWidth = 10;
                        MySheet.Cells[lastRow, 3].EntireRow.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                        MySheet.Cells[lastRow, 4].EntireColumn.ColumnWidth = 5;
                        lastRow += 1;
                    }
                    Excel.Range styling = MySheet.get_Range("E:E", System.Type.Missing);
                    styling.EntireColumn.ColumnWidth = 60;
                }

                DateTime now = DateTime.Now;
                //DateTime yesterday = DateTime.Now.AddDays(-1);
                DateTime yesterday = DateTime.Now.AddDays(-1);

                foreach (Pool pool in SelectedPools)
                {
                    DayOfWeek[] week = new[] { DayOfWeek.Sunday, DayOfWeek.Monday, DayOfWeek.Tuesday,
                                     DayOfWeek.Wednesday,DayOfWeek.Thursday,DayOfWeek.Friday,
                                     DayOfWeek.Saturday};
                    if (pool.BackupDay == 7)
                    {
                        yesterday = DateTime.Now.AddDays(-1);
                    }
                    else
                    {
                        for (int day = -7; day < 0; day++)
                        {
                            yesterday = DateTime.Now.AddDays(day);
                            if (yesterday.DayOfWeek == week[pool.BackupDay])
                            {
                                break;
                            }
                        }
                    }

                    ICollection<BackupServer> servers = db.BackupServers.Where(id => id.Pool.Id == pool.Id).ToArray();

                    foreach (BackupServer server in servers)
                    {
                        int row_index = 5;
                        Excel.Range range = MySheet.get_Range("B1",
                            "Z" + MySheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row);

                        Dictionary<string, string> result = CheckBackupServer(server, yesterday, range);

                        BackupServer_Report serverReport = RegisterResult(report, server, result, false);
                        if (autobesr && serverReport.State != "OK")
                        {
                            if (serverReport.Ping != "Ping Success")
                            {
                                serverReport.Services = "Services: Ping KO";
                                serverReport.Relaunched = "Non relancées: Ping KO";
                            }
                            else
                            {
                                string services = ServiceLauncher(server.Id);
                                string BESROK = "Backup Exec System Recovery: Running: Manual";
                                string SYMOK = "SymSnapService: Running: Manual";
                                if ((services.IndexOf(BESROK, StringComparison.OrdinalIgnoreCase) > 0) &&
                                    services.IndexOf(SYMOK, StringComparison.OrdinalIgnoreCase) > 0)
                                {
                                    serverReport.Services = "OK";
                                    string relaunched = BackupExecLauncher(server.Id);
                                    if (relaunched.IndexOf("Sauvegarde OK") != -1)
                                    {
                                        serverReport.Relaunched = "Relancées";
                                    }
                                }
                                else
                                {
                                    serverReport.Services = "KO";
                                    serverReport.Relaunched = "Non relancées: Services non lancés.";
                                }
                            }
                            if (ModelState.IsValid)
                            {
                                db.Entry(serverReport).State = System.Data.Entity.EntityState.Modified;
                                db.SaveChanges();
                            }
                        }
                        UpdateCells(MySheet, server.Name, serverReport, row_index);
                    }
                }

            CHECK_FINALIZER:
                {
                    try
                    {
                        MyWorkbook.SaveAs(HomeController.BESR_RESULTS_FOLDER + "Controle_hebdomadaire_sauvegarde_BESR_PCI_" + DateTime.Now.ToString("yyyy") + "-" +
                            report.Id.ToString() + ".xlsx",
                            Excel.XlFileFormat.xlOpenXMLWorkbook, null, null, null, null,
                            Excel.XlSaveAsAccessMode.xlNoChange, null, null, null, null, null);

                        report.ResultPath = HomeController.BESR_RESULTS_FOLDER + "Controle_hebdomadaire_sauvegarde_BESR_PCI_" + DateTime.Now.ToString("yyyy") + "-" +
                            report.Id.ToString() + ".xlsx";
                    }
                    catch (Exception saveException)
                    {
                        Specific_Logging(saveException, "BesrChecker");
                        ExecutionErrors += "Erreur de sauvegarde: " + saveException.Message + "\r\n";
                        try
                        {
                            MyWorkbook.SaveAs(HomeController.BESR_RESULTS_FOLDER + "Controle_hebdomadaire_sauvegarde_BESR_PCI_" + DateTime.Now.ToString("yyyy") + ".xlsx",
                                Excel.XlFileFormat.xlOpenXMLWorkbook, null, null, null, null,
                                Excel.XlSaveAsAccessMode.xlNoChange, null, null, null, null, null);
                            report.ResultPath = HomeController.BESR_RESULTS_FOLDER + "Controle_hebdomadaire_sauvegarde_BESR_PCI_" + DateTime.Now.ToString("yyyy") + ".xlsx";
                        }
                        catch { }
                    }
                    report.Duration = DateTime.Now.Subtract(report.DateTime);
                    report.LastUpdate = DateTime.Now;
                    report.TotalChecked = report.BackupServer_Reports.Count;
                    report.TotalErrors = report.BackupServer_Reports.Where(serverreport => serverreport.State != "OK").Count();
                    if (ModelState.IsValid)
                    {
                        db.Entry(report).State = System.Data.Entity.EntityState.Modified;
                        db.SaveChanges();
                        string buildOk = BuildBackupEmail(email.Id);
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
            }
            catch (Exception running)
            {
                Specific_Logging(running, "BesrChecker");
            }
            finally
            {
                McoUtilities.CloseExcel(MyApplication, MyWorkbook, MySheet);
            }
            results["response"] = "OK";
            results["email"] = emailId.ToString();
            results["errors"] = "Fin d'exécution. \n" + "Erreurs d'exécution : " + ExecutionErrors;
            Specific_Logging(new Exception("...."), "BesrChecker", 3);
            return Json(results, JsonRequestBehavior.AllowGet);
        }

        public JsonResult CheckPool(int id)
        {
            UpdatePoolDatabase();
            Dictionary<string, string> results = new Dictionary<string, string>();
            results["response"] = "";
            results["email"] = "";
            results["errors"] = "";
            int emailId = 0;
            string ExecutionErrors = "";
            try
            {
                MyApplication = new Excel.Application();
                MyApplication.Visible = false;
                MyApplication.DisplayAlerts = false;
                bool foundedFile = false;
                string[] report_files = System.IO.Directory.GetFiles(HomeController.BESR_RESULTS_FOLDER,
                    "Controle_hebdomadaire_sauvegarde_BESR_PCI_" + DateTime.Now.ToString("yyyy") + "*.xlsx", System.IO.SearchOption.TopDirectoryOnly);
                string filename = "Controle_hebdomadaire_sauvegarde_BESR_PCI_" + DateTime.Now.ToString("yyyy") + ".xlsx";
                if (report_files.Length > 0)
                {
                    filename = Directory.GetFiles(HomeController.BESR_RESULTS_FOLDER,
                        "Controle_hebdomadaire_sauvegarde_BESR_PCI_" + DateTime.Now.ToString("yyyy") + "*.xlsx", System.IO.SearchOption.TopDirectoryOnly)
                            .Select(x => new FileInfo(x))
                            .OrderByDescending(x => x.LastWriteTime).FirstOrDefault().FullName;

                    MyWorkbook = MyApplication.Workbooks.Open(filename);
                    foundedFile = true;
                }
                else
                {
                    ExecutionErrors += HomeController.BESR_RESULTS_FOLDER + "Controle_hebdomadaire_sauvegarde_BESR_PCI_" + DateTime.Now.ToString("yyyy") + "-(id).xlsx n'a pas été trouvé.\r\n";
                    MyWorkbook = MyApplication.Workbooks.Add(Excel.XlWBATemplate.xlWBATWorksheet);
                    foundedFile = false;
                }

                CalendarWeekRule weekRule = CalendarWeekRule.FirstDay;
                DayOfWeek firstWeekDay = DayOfWeek.Monday;
                Calendar calendar = System.Threading.Thread.CurrentThread.CurrentCulture.Calendar;
                int currentWeek = calendar.GetWeekOfYear(DateTime.Now, weekRule, firstWeekDay);

                string sheetName = "Semaine " + currentWeek.ToString();
                bool foundedSheet = false;
                foreach (Excel.Worksheet sheet in MyWorkbook.Sheets)
                {
                    if (sheet.Name.StartsWith(sheetName))
                    {
                        sheetName = sheet.Name;
                        foundedSheet = true;
                        break;
                    }
                }
                if (foundedSheet)
                {
                    MySheet = (Excel.Worksheet)MyWorkbook.Sheets[sheetName];
                }
                else
                {
                    MySheet = (Excel.Worksheet)MyWorkbook.Sheets.Add();
                    MySheet.Name = sheetName + " " + DateTime.Now.ToString("dd") +
                        DateTime.Now.ToString("MM") + DateTime.Now.ToString("yyyy");
                }

                MySheet.Activate();

                //Database objects init
                IQueryable<BackupReport> oldReports = db.BackupReports.Where(reportid => reportid.WeekNumber == currentWeek); //.First();
                BackupReport report;
                Email email;

                if (oldReports.Count() == 0)
                {
                    report = db.BackupReports.Create();
                    report.DateTime = DateTime.Now;
                    report.LastUpdate = DateTime.Now;
                    report.WeekNumber = currentWeek;
                    report.TotalChecked = 0;
                    report.TotalErrors = 0;
                    report.ResultPath = "";
                    report.Module = HomeController.BESR_MODULE;
                    report.Author = User.Identity.Name;

                    email = db.Emails.Create();
                    report.Email = email;
                    email.Report = report;
                    email.Module = HomeController.BESR_MODULE;
                    email.Recipients = "";
                    email = Emails_Controller.SetRecipients(email, HomeController.BESR_MODULE);
                    if (ModelState.IsValid)
                    {
                        db.BackupReports.Add(report);
                        db.SaveChanges();
                        emailId = report.Email.Id;
                        int reportNumber = db.BackupReports.Count();
                        if (reportNumber > HomeController.BESR_MAX_REPORT_NUMBER)
                        {
                            int reportNumberToDelete = reportNumber - HomeController.BESR_MAX_REPORT_NUMBER;
                            BackupReport[] reportsToDelete =
                                db.BackupReports.OrderBy(idReport => idReport.Id).Take(reportNumberToDelete).ToArray();
                            foreach (BackupReport toDeleteReport in reportsToDelete)
                            {
                                DeleteBackupReport(toDeleteReport.Id);
                            }
                        }
                    }
                    else
                    {
                        results["response"] = "KO";
                        results["email"] = null;
                        results["errors"] = "Impossible de créer un rapport dans la base de données.";
                        goto CHECK_FINALIZER;
                    }
                }
                else
                {
                    report = oldReports.First();
                    report.DateTime = DateTime.Now;
                    email = report.Email;
                    email.Recipients = "";
                    email = Emails_Controller.SetRecipients(email, HomeController.BESR_MODULE);
                    if (ModelState.IsValid)
                    {
                        db.Entry(report).State = System.Data.Entity.EntityState.Modified;
                        db.SaveChanges();
                        emailId = report.Email.Id;
                    }
                    else
                    {
                        results["response"] = "KO";
                        results["email"] = null;
                        results["errors"] = "Impossible de créer un rapport dans la base de données.";
                        goto CHECK_FINALIZER;
                    }
                }

                //End database objects init

                int lastRow = MySheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row;
                ICollection<BackupServer> list = db.BackupServers.OrderBy(idServer => idServer.Pool.Id).ToArray();
                if (!foundedSheet)
                {
                    foreach (BackupServer server in list)
                    {
                        Color rangeColor = ContrastColor(server.Pool.CellColor);
                        Excel.Range ActualRange = MySheet.get_Range("A" + lastRow,
                                "M" + lastRow);
                        ActualRange.Interior.Color = System.Drawing.ColorTranslator.FromHtml(server.Pool.CellColor);
                        ActualRange.Font.Color = System.Drawing.ColorTranslator.FromHtml(ColorTranslator.ToHtml(rangeColor));
                        MySheet.Cells[lastRow, 1] = server.Pool.Name;
                        MySheet.Cells[lastRow, 1].EntireColumn.ColumnWidth = 15;
                        MySheet.Cells[lastRow, 2] = server.Name;
                        MySheet.Cells[lastRow, 2].EntireColumn.ColumnWidth = 15;
                        MySheet.Cells[lastRow, 3] = server.Disks;
                        MySheet.Cells[lastRow, 3].EntireColumn.ColumnWidth = 10;
                        MySheet.Cells[lastRow, 3].EntireRow.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                        MySheet.Cells[lastRow, 4].EntireColumn.ColumnWidth = 5;
                        lastRow += 1;
                    }
                    Excel.Range styling = MySheet.get_Range("E:E", System.Type.Missing);
                    styling.EntireColumn.ColumnWidth = 60;
                }

                DateTime now = DateTime.Now;
                //DateTime yesterday = DateTime.Now.AddDays(-1);
                DateTime yesterday = DateTime.Now.AddDays(-1);
                int dayOfWeek = (int)now.Date.DayOfWeek;

                Pool pool = db.Pools.Find(id);
                DayOfWeek[] week = new[] { DayOfWeek.Sunday, DayOfWeek.Monday, DayOfWeek.Tuesday,
                                     DayOfWeek.Wednesday,DayOfWeek.Thursday,DayOfWeek.Friday,
                                     DayOfWeek.Saturday};
                if (pool.BackupDay == 7)
                {
                    yesterday = DateTime.Now.AddDays(-1);
                }
                else
                {
                    for (int day = -7; day < 0; day++)
                    {
                        yesterday = DateTime.Now.AddDays(day);
                        if (yesterday.DayOfWeek == week[pool.BackupDay])
                        {
                            break;
                        }
                    }
                }
                ICollection<BackupServer> servers = db.BackupServers.Where(idServer => idServer.Pool.Id == pool.Id).ToArray();
                foreach (BackupServer server in servers)
                {
                    int row_index = 5;
                    Excel.Range range = MySheet.get_Range("B1",
                        "Z" + MySheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row);

                    Dictionary<string, string> result = CheckBackupServer(server, yesterday, range);

                    BackupServer_Report serverReport = RegisterResult(report, server, result, false);
                    UpdateCells(MySheet, server.Name, serverReport, row_index);
                }
            CHECK_FINALIZER:
                {
                    try
                    {
                        MyWorkbook.SaveAs(HomeController.BESR_RESULTS_FOLDER + "Controle_hebdomadaire_sauvegarde_BESR_PCI_" + DateTime.Now.ToString("yyyy") + "-" +
                            report.Id.ToString() + ".xlsx",
                            Excel.XlFileFormat.xlOpenXMLWorkbook, null, null, null, null,
                            Excel.XlSaveAsAccessMode.xlNoChange, null, null, null, null, null);

                        report.ResultPath = HomeController.BESR_RESULTS_FOLDER + "Controle_hebdomadaire_sauvegarde_BESR_PCI_" + DateTime.Now.ToString("yyyy") + "-" +
                            report.Id.ToString() + ".xlsx";
                    }
                    catch (Exception saveException)
                    {
                        Specific_Logging(saveException, "CheckPool");
                        ExecutionErrors += "Erreur de sauvegarde: " + saveException.Message + "\r\n";
                        try
                        {
                            MyWorkbook.SaveAs(HomeController.BESR_RESULTS_FOLDER + "Controle_hebdomadaire_sauvegarde_BESR_PCI_" + DateTime.Now.ToString("yyyy") + ".xlsx",
                                Excel.XlFileFormat.xlOpenXMLWorkbook, null, null, null, null,
                                Excel.XlSaveAsAccessMode.xlNoChange, null, null, null, null, null);
                            report.ResultPath = HomeController.BESR_RESULTS_FOLDER + "Controle_hebdomadaire_sauvegarde_BESR_PCI_" + DateTime.Now.ToString("yyyy") + ".xlsx";
                        }
                        catch { }
                    }
                    report.Duration = DateTime.Now.Subtract(report.DateTime);
                    report.LastUpdate = DateTime.Now;
                    report.TotalChecked = report.BackupServer_Reports.Count;
                    report.TotalErrors = report.BackupServer_Reports.Where(serverreport => serverreport.State != "OK").Count();
                    if (ModelState.IsValid)
                    {
                        db.Entry(report).State = System.Data.Entity.EntityState.Modified;
                        db.SaveChanges();
                        string buildOk = BuildBackupEmail(email.Id);
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
            }
            catch (Exception running)
            {
                Specific_Logging(running, "CheckPool");
            }
            finally
            {
                McoUtilities.CloseExcel(MyApplication, MyWorkbook, MySheet);
            }
            results["response"] = "OK";
            results["email"] = emailId.ToString();
            results["errors"] = "Fin d'exécution. \n" + "Erreurs d'exécution : " + ExecutionErrors;
            Specific_Logging(new Exception("...."), "CheckPool", 3);
            return Json(results, JsonRequestBehavior.AllowGet);
        }

        public Dictionary<string, string> CheckBackupFileDate(BackupServer server, string file, DateTime yesterday)
        {
            Dictionary<string, string> result = new Dictionary<string, string>();
            result.Add("state", "");
            result.Add("details", "");
            try
            {
                DateTime lastupdate = System.IO.File.GetLastWriteTime(file);
                if (yesterday.Date.CompareTo(lastupdate.Date) <= 0)
                {
                    result["state"] = "OK";
                }
                else
                {
                    result["state"] = "KO";
                    result["details"] = "Date non valide";
                }
            }
            catch (Exception exception)
            {
                result["state"] = "KO";
                result["details"] = exception.Message;
            }
            return result;
        }

        public Dictionary<string, string> CheckBackupFilesNumber(BackupServer server, DateTime yesterday, string folder_path, string[] disks)
        {
            Dictionary<string, string> result = new Dictionary<string, string>();
            result.Add("state", "");
            result.Add("details", "");
            try
            {
                string[] files = System.IO.Directory.GetFiles(folder_path, "*.*v2i");
                if ((files == null) || (files.Count() < (disks.Count() + 1)))
                {
                    result["state"] = "KO";
                    result["details"] = "Nombre de fichiers trop petit.";
                }
                else
                {
                    if (files.Count() > (disks.Count() + 1))
                    {
                        bool testfiles = GoodFilesKepper(server, yesterday);
                        if (!testfiles)
                        {
                            result["state"] = "KO";
                            result["details"] = "Nombre de fichiers trop grand.";
                        }
                        else
                        {
                            result["state"] = "OK";
                        }
                    }
                    else
                    {
                        result["state"] = "OK";
                    }
                }
            }
            catch (Exception exception)
            {
                result["state"] = "KO";
                result["details"] = exception.Message;
            }
            return result;
        }

        public Dictionary<string, string> CheckBackupIndexFileNumber(BackupServer server, string folder_path)
        {
            Dictionary<string, string> result = new Dictionary<string, string>();
            result.Add("state", "");
            result.Add("details", "");
            try
            {
                string[] indexFiles = System.IO.Directory.GetFiles(folder_path, "*.sv2i");
                if (indexFiles == null || indexFiles.Count() < 1)
                {
                    result["state"] = "KO";
                    result["details"] = "Fichier d'indexation manquant. ";
                }
                else
                {
                    if (indexFiles.Count() > 1)
                    {
                        result["state"] = "KO";
                        result["details"] = "Plusieurs fichiers d'indexation. ";
                    }
                    else
                    {
                        result["state"] = "OK";
                    }
                }
            }
            catch (Exception exception)
            {
                result["state"] = "KO";
                result["details"] = exception.Message;
            }
            return result;
        }

        public Dictionary<string, string> CheckBackupServerDisks(BackupServer server, DateTime yesterday, string[] disks, string[] files)
        {
            Dictionary<string, string> result = new Dictionary<string, string>();
            result.Add("state", "");
            result.Add("details", "");
            try
            {
                foreach (string disk in disks)
                {
                    if (disk.Trim() == " " || disk.Trim() == "")
                    {
                        continue;
                    }
                    bool found = false;
                    foreach (string file in files)
                    {
                        if (file.IndexOf(server.Name, StringComparison.OrdinalIgnoreCase) > 0
                                        && file.IndexOf("_" + disk.Trim() + "_", StringComparison.OrdinalIgnoreCase) > 0
                                        && file.IndexOf(".v2i", StringComparison.OrdinalIgnoreCase) > 0)
                        {
                            found = true;
                            result = CheckBackupFileDate(server, file, yesterday);
                            break;
                        }
                        else
                        {
                            continue;
                        }
                    }
                    if (found && result["state"] == "OK")
                    {
                        continue;
                    }
                    else
                    {
                        if (!found)
                        {
                            result["state"] = "KO";
                            result["details"] = "Sauvegarde disque " + disk + " non trouvée";
                        }
                        break;
                    }
                }
            }
            catch (DirectoryNotFoundException)
            {
                result["state"] = "KO";
                result["details"] = "Répertoire Absent. ";
            }
            catch (UnauthorizedAccessException)
            {
                result["state"] = "KO";
                result["details"] = "Accès refusé. ";
            }
            catch (Exception)
            {
                result["state"] = "KO";
                result["details"] = "Erreur inconue. ";
            }
            return result;
        }

        public Account GetPoolAccount(int id = 0)
        {
            Account account = new Account();
            if (id != 0)
            {
                Pool pool = db.Pools.Find(id);
                if (pool != null)
                {
                    account = db.Accounts.Where(acc => acc.DisplayName.ToUpper() == pool.CheckAccount.ToUpper())
                        .FirstOrDefault();
                }
            }
            if (account == null || account.DisplayName == null || account.DisplayName.Trim() == "")
            {
                account = new Account();
                account.DisplayName = HomeController.DEFAULT_DOMAIN_AND_USERNAME_IMPERSONNATION;
                account.Domain = HomeController.DEFAULT_DOMAIN_IMPERSONNATION;
                account.Username = HomeController.DEFAULT_USERNAME_IMPERSONNATION;
                account.Password = HomeController.DEFAULT_PASSWORD_IMPERSONNATION;
            }
            return account;
        }

        public string GetFolderPath(BackupServer server)
        {
            string folder_path = HomeController.BACKUP_REMOTE_CHECK_SERVER_ROOT_FOLDER + server.Pool.Name + "\\" +
                                server.Name;
            if (server.Pool.CheckFolder != null && server.Pool.CheckFolder.Trim() != "")
            {
                folder_path = server.Pool.CheckFolder + server.Name;
            }
            return folder_path;
        }

        public Dictionary<string, string> CheckBackupServer(BackupServer server, DateTime yesterday, Excel.Range range, Account account = null)
        {
            if (account == null)
            {
                account = GetPoolAccount(server.PoolId);
            }
            Dictionary<string, string> result = new Dictionary<string, string>();
            result.Add("state", "");
            result.Add("details", "");
            int row_index = 5;
            try
            {
                string folder_Path = GetFolderPath(server);

                string[] disks = server.Disks.Split(',');
                using (UNC_ACCESSOR)
                {
                    UNC_ACCESSOR.NetUseWithCredentials(folder_Path,
                        account.Username, account.Domain,
                        McoUtilities.Decrypt(account.Password));
                    string[] files = System.IO.Directory.GetFiles(folder_Path, "*.*v2i");

                    //CHECKING ALL FILES NUMBER
                    result = CheckBackupFilesNumber(server, yesterday, folder_Path, disks);
                    if (!IsValidResult(result))
                    {
                        return result;
                    }
                    //CHECKING BACKUP FILES DATE
                    result = CheckBackupServerDisks(server, yesterday, disks, files);
                    if (!IsValidResult(result))
                    {
                        return result;
                    }
                    //CHECKING INDEX FILES
                    result = CheckBackupIndexFileNumber(server, folder_Path);
                    if (!IsValidResult(result))
                    {
                        return result;
                    }
                    string[] indexFiles = System.IO.Directory.GetFiles(folder_Path, "*.sv2i");
                    //CHECKING INDEX FILES
                    result = CheckBackupFileDate(server, indexFiles[0], yesterday);
                    if (!IsValidResult(result))
                    {
                        return result;
                    }
                }
            }
            catch (DirectoryNotFoundException)
            {
                result["state"] = "KO";
                result["details"] = "Répertoire Absent. ";
            }
            catch (UnauthorizedAccessException)
            {
                result["state"] = "KO";
                result["details"] = "Accès refusé. ";
            }
            catch (Exception)
            {
                result["state"] = "KO";
                result["details"] = "Erreur inconue. ";
            }
            return result;
        }

        public bool IsValidResult(Dictionary<string, string> result)
        {
            try
            {
                if (result["state"] == "OK")
                {
                    return true;
                }
            }
            catch { }
            return false;
        }

        public bool UpdateCells(Excel.Worksheet sheet, string servername, BackupServer_Report serverReport, int row_index)
        {
            try
            {
                Excel.Range range = MySheet.get_Range("B1",
                       "B" + sheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row);
                foreach (Excel.Range cell in range)
                {
                    if (cell.Value != null && cell.Value == servername)
                    {
                        sheet.Cells[cell.Row, 4] = serverReport.State;
                        sheet.Cells[cell.Row, 4].Interior.Color = System.Drawing.ColorTranslator.FromHtml("#fff");
                        sheet.Cells[cell.Row, 4].Font.Color = System.Drawing.ColorTranslator.FromHtml(GetMathingColor(serverReport.State));
                        sheet.Cells[cell.Row, 5] = serverReport.Details;
                        sheet.Cells[cell.Row, 6] = serverReport.Ping;
                        sheet.Cells[cell.Row, 7] = serverReport.Relaunched;
                        if (serverReport.State == "OK")
                        {
                            string folder_Path = GetFolderPath(serverReport.BackupServer);
                            string[] files = null;
                            using (UNC_ACCESSOR)
                            {
                                UNC_ACCESSOR.NetUseWithCredentials(folder_Path,
                                    HomeController.DEFAULT_USERNAME_IMPERSONNATION,
                                    HomeController.DEFAULT_DOMAIN_IMPERSONNATION,
                                    McoUtilities.Decrypt(HomeController.DEFAULT_PASSWORD_IMPERSONNATION));
                                files = System.IO.Directory.GetFiles(folder_Path, "*.*v2i");
                            }
                            foreach (string file in files)
                            {
                                if (file.Trim() == " " || file.Trim() == "")
                                {
                                    continue;
                                }
                                DateTime lastupdate = System.IO.File.GetLastWriteTime(file);
                                sheet.Cells[cell.Row, row_index] = file;
                                sheet.Cells[cell.Row, row_index].EntireColumn.ColumnWidth = 60;
                                sheet.Cells[cell.Row, row_index + 1] = lastupdate.ToString();
                                sheet.Cells[cell.Row, row_index + 1].EntireColumn.ColumnWidth = 25;
                                row_index = row_index + 2;
                            }
                        }
                        break;
                    }
                }
                return true;
            }
            catch (Exception exception)
            {
                Specific_Logging(exception, "UpdateCells");
                return false;
            }
        }

        public BackupServer_Report RegisterResult(BackupReport report, BackupServer server, Dictionary<string, string> result, bool relaunch)
        {
            try
            {
                //Empty the old report on the server
                ICollection<BackupServer_Report> oldserverreports = db.BackupServerReports.Where(
                        serverreportid => serverreportid.BackupServerId == server.Id).ToList(); //.First();
                foreach (BackupServer_Report oldserverreport in oldserverreports)
                {
                    if (oldserverreport.BackupReport.WeekNumber == report.WeekNumber)
                    {
                        db.BackupServerReports.Remove(oldserverreport);
                        db.SaveChanges();
                    }
                }
                BackupServer_Report serverReport = new BackupServer_Report();
                serverReport.BackupReport = report;
                serverReport.BackupServer = server;
                serverReport.Details = serverReport.Ping =
                    serverReport.Services = serverReport.Relaunched = "";
                serverReport.State = "KO";
                serverReport.State = result["state"];
                serverReport.Details = result["details"];
                //SERVER N-OK
                if (!IsValidResult(result))
                {
                    try
                    {
                        Ping ping = new Ping();
                        PingOptions options = new PingOptions(64, true);
                        PingReply pingreply = ping.Send(server.Name);
                        serverReport.Ping = "Ping " + pingreply.Status.ToString();
                    }
                    catch
                    {
                        serverReport.Ping = "Ping KO";
                        serverReport.Services = "Services: Ping KO";
                        serverReport.Relaunched = "Non relancées: Ping KO";
                    }
                }
                //SERVER OK
                else
                {
                    serverReport.Ping = "Ping OK";
                    serverReport.Services = "";
                    serverReport.Relaunched = "";
                }
                if (ModelState.IsValid)
                {
                    db.BackupServerReports.Add(serverReport);
                    db.SaveChanges();
                }
                return serverReport;
            }
            catch (Exception exception)
            {
                Specific_Logging(exception, "RegisterResult");
            }
            return null;
        }

        public string GetMathingColor(string state)
        {
            if (state == "OK")
            {
                return "#00b050";
            }
            else
            {
                return "#ff0000";
            }
        }

        public string BuildBackupEmail(int id)
        {
            Email email = db.Emails.Find(id);
            BackupReport report = (BackupReport)email.Report;
            if (email == null)
            {
                return HttpNotFound().ToString();
            }

            string body = "<style>tr:hover>td{cursor:pointer;background-color:#68b3ff;}</style><div><br /><span>Rapport de check BESR.</span><br />" +
                "Ci dessous, un tableau récapitulatif de l'etat des sauvegardes Windows pour l'ensemble des serveurs testés. " +
                "Pour les serveurs d'un Pool donné, la colonne 'Statut' est vide lorsque le check n'a pas encore été exécuté " +
                    "au cours de la semaine " + report.WeekNumber + ".<br />";
            body += "<table style='position:relative;width:100%;background-color :#aca8a4;'></table>" +
                "Nombre total de serveurs vérifiés : " + email.Report.TotalChecked + "<br />" +
                "Nombre total de serveurs en erreur : " + email.Report.TotalErrors + "<br />";
            double percentage = Math.Round((double)email.Report.TotalErrors / (double)email.Report.TotalChecked, 4) * 100;
            body += "<span style='font-weight:bold;color:#e75114;'>Pourcentage de serveurs en erreurs : " + percentage + "%</span><br />";
            body += "Date de dernière mise à jour : " + DateTime.Now.ToString() + "<br />";

            body += "<table style='position:relative;width:100%;background-color :#aca8a4;'><thead></thead><tbody>";
            BackupServer_Report[] backupserverreports = db.BackupServerReports.Where(
                serverreportid => serverreportid.BackupReport.Id == email.Report.Id)
                .OrderBy(pool => pool.BackupServer.Pool.Id).ToArray();
            List<Pool> pools = new List<Pool>();
            foreach (BackupServer_Report serverreport in backupserverreports.Reverse())
            {
                if (!pools.Contains(serverreport.BackupServer.Pool))
                {
                    pools.Add(serverreport.BackupServer.Pool);
                    body += "</tbody></table><br />";
                    body += "<table style='position:relative;width:100%;background-color :#aca8a4;'><thead></thead><tbody></tbody></table>";
                    body += serverreport.BackupServer.Pool.Name.ToUpper() + "<br />";
                    int serverchecked = serverreport.BackupServer.Pool.BackupServers.Count();
                    int serverfailed = db.BackupServerReports.Where(
                        serverreportid => serverreportid.BackupReport.Id == email.Report.Id)
                        .Where(serverreportid => serverreportid.BackupServer.PoolId == serverreport.BackupServer.Pool.Id).Where(serverstate => serverstate.State != "OK").Count();
                    body += "Nombre de serveurs dans le Pool : " + serverchecked.ToString() + "<br />" +
                    "Nombre de serveurs en erreur : " + serverfailed.ToString() + "<br />" +
                    "<span style='font-weight:bold;color:#e75114;'>Pourcentage de serveurs en erreur: " + (Math.Round((double)serverfailed / (double)serverchecked, 4) * 100).ToString() + "%</span>" +
                    "<table style='position:relative;width:100%;background-color :#aca8a4;'></table>";
                    body += "<table style='position:relative;width:100%;'><thead>" +
                    "<tr style='font-weight:bold;position:relative;width:100%;height:35px;background-color:#ecebeb'><th>Pools</th><th>Serveurs</th><th>Partitions</th><th>Statut</th>" +
                    "<th>Erreurs</th><th>Ping</th><th>Sauvegardes</th></thead><tbody>";
                }
                string okayColor = "", pingColor = "";
                Color font = ContrastColor(serverreport.BackupServer.Pool.CellColor);
                if (serverreport.State == "OK")
                {
                    okayColor = "#00b050";
                }
                else
                {
                    okayColor = "#ff0000";
                }
                string ping = (serverreport.Ping == "Ping OK" || serverreport.Ping == "Ping Success") ? "Ping OK" : "Ping KO";
                if (ping == "Ping OK")
                {
                    pingColor = "";
                }
                else
                {
                    pingColor = "background-color:#fff;color:#ff0000";
                }
                body += "<tr style='position:relative;text-align:center;font-weight:bold;color:" + ColorTranslator.ToHtml(font) + ";width:100%;min-height:40px;border:1px solid #000;background-color:" + serverreport.BackupServer.Pool.CellColor + ";'>" +
                        "<td style='position:relative;text-align:center;width:100px;border:5px solid " + serverreport.BackupServer.Pool.CellColor + ";'>" + serverreport.BackupServer.Pool.Name + "</td>" +
                        "<td style='position:relative;text-align:center;width:100px;border:5px solid " + serverreport.BackupServer.Pool.CellColor + ";'>" + serverreport.BackupServer.Name + "</td>" +
                        "<td style='position:relative;text-align:center;width:80px;border:5px solid " + serverreport.BackupServer.Pool.CellColor + ";'>" + serverreport.BackupServer.Disks + "</td>" +
                        "<td style='position:relative;text-align:center;width:40px;border:5px solid " + serverreport.BackupServer.Pool.CellColor + ";background-color:#fff;color:" + okayColor + "'>" + serverreport.State + "</td>";
                body += "<td style='position:relative;text-align:center;width:80px;border:5px solid " + serverreport.BackupServer.Pool.CellColor + ";'>" + serverreport.Details + "</td>";
                body += "<td style='position:relative;text-align:center;width:60px;border:5px solid " + serverreport.BackupServer.Pool.CellColor + ";" + pingColor + "'>" + ping + "</td>";
                body += "<td style='position:relative;text-align:center;width:100px;border:5px solid " + serverreport.BackupServer.Pool.CellColor + ";'>" + serverreport.Relaunched + "</td></tr>";
            }
            body += "</tbody></table>";
            email.Body = body;
            email.Subject = "Resultat check BESR du " + email.Report.DateTime.ToString();
            if (ModelState.IsValid)
            {
                db.Entry(email).State = System.Data.Entity.EntityState.Modified;
                db.SaveChanges();
            }
            return "BuildOK";
        }

        public string UpdateBackupServerFile(BackupServer server, string modified)
        {
            string response = "KO";
            try
            {
                MyApplication = new Excel.Application();
                MyApplication.Visible = false;
                MyApplication.DisplayAlerts = false;
                string[] report_files = System.IO.Directory.GetFiles(HomeController.BESR_RESULTS_FOLDER,
                    "Controle_hebdomadaire_sauvegarde_BESR_PCI_" + DateTime.Now.ToString("yyyy") + "*.xlsx", System.IO.SearchOption.TopDirectoryOnly);
                string filename = "Controle_hebdomadaire_sauvegarde_BESR_PCI_" + DateTime.Now.ToString("yyyy") + ".xlsx";
                if (report_files.Length > 0)
                {
                    filename = Directory.GetFiles(HomeController.BESR_RESULTS_FOLDER,
                        "Controle_hebdomadaire_sauvegarde_BESR_PCI_" + DateTime.Now.ToString("yyyy") + "*.xlsx", System.IO.SearchOption.TopDirectoryOnly)
                            .Select(x => new FileInfo(x))
                            .OrderByDescending(x => x.LastWriteTime).FirstOrDefault().FullName;

                    MyWorkbook = MyApplication.Workbooks.Open(filename);
                    CalendarWeekRule weekRule = CalendarWeekRule.FirstDay;
                    DayOfWeek firstWeekDay = DayOfWeek.Monday;
                    Calendar calendar = System.Threading.Thread.CurrentThread.CurrentCulture.Calendar;
                    int currentWeek = calendar.GetWeekOfYear(DateTime.Now, weekRule, firstWeekDay);

                    string sheetName = "Semaine " + currentWeek.ToString();
                    bool foundedSheet = false;
                    foreach (Excel.Worksheet sheet in MyWorkbook.Sheets)
                    {
                        if (sheet.Name.StartsWith(sheetName))
                        {
                            sheetName = sheet.Name;
                            foundedSheet = true;
                            break;
                        }
                    }
                    if (foundedSheet && modified == "server")
                    {
                        MySheet = (Excel.Worksheet)MyWorkbook.Sheets[sheetName];
                        MySheet.Activate();
                        int lastRow = MySheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row;
                        lastRow++;
                        Excel.Range ActualRange = MySheet.get_Range("A" + lastRow,
                        "M" + lastRow);
                        ActualRange.Interior.Color = System.Drawing.ColorTranslator.FromHtml(server.Pool.CellColor);
                        MySheet.Cells[lastRow, 1] = server.Pool.Name;
                        MySheet.Cells[lastRow, 1].EntireColumn.ColumnWidth = 15;
                        MySheet.Cells[lastRow, 2] = server.Name;
                        MySheet.Cells[lastRow, 2].EntireColumn.ColumnWidth = 15;
                        MySheet.Cells[lastRow, 3] = server.Disks;
                        MySheet.Cells[lastRow, 3].EntireColumn.ColumnWidth = 10;
                        MySheet.Cells[lastRow, 3].EntireRow.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                        MySheet.Cells[lastRow, 4].EntireColumn.ColumnWidth = 5;
                        Excel.Range styling = MySheet.get_Range("E:E", System.Type.Missing);
                        styling.EntireColumn.ColumnWidth = 60;
                        try
                        {
                            MyWorkbook.SaveAs(filename,
                                Excel.XlFileFormat.xlOpenXMLWorkbook, null, null, null, null,
                                Excel.XlSaveAsAccessMode.xlNoChange, null, null, null, null, null);
                            response = "OK";
                        }
                        catch (Exception exception)
                        {
                            response = exception.Message;
                        }
                    }
                    if (foundedSheet && modified == "disk")
                    {
                        MySheet = (Excel.Worksheet)MyWorkbook.Sheets[sheetName];
                        MySheet.Activate();
                        Excel.Range range = MySheet.get_Range("A1",
                            "C" + MySheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row);

                        foreach (Excel.Range cell in range)
                        {
                            if (cell.Value == server.Name)
                            {
                                MySheet.Cells[cell.Row, 3] = server.Disks;
                                break;
                            }
                        }
                        try
                        {
                            MyWorkbook.SaveAs(filename,
                                Excel.XlFileFormat.xlOpenXMLWorkbook, null, null, null, null,
                                Excel.XlSaveAsAccessMode.xlNoChange, null, null, null, null, null);
                            response = "OK";
                        }
                        catch (Exception exception)
                        {
                            response = exception.Message;
                        }
                    }
                }
            }
            catch { }
            finally
            {
                McoUtilities.CloseExcel(MyApplication, MyWorkbook, MySheet);
            }
            return response;
        }

        [HttpPost]
        public string AddPool()
        {
            try
            {
                string poolname = Request.Form["poolname"];
                string backupmanager = Request.Form["backupmanager"];
                string checkfolder = Request.Form["checkfolder"];
                if (backupmanager == null || backupmanager.Trim() == "")
                {
                    backupmanager = HomeController.BACKUP_REMOTE_SERVER_EXEC;
                }
                if (checkfolder == null || checkfolder.Trim() == "")
                {
                    checkfolder = HomeController.BACKUP_REMOTE_CHECK_SERVER_ROOT_FOLDER + poolname.ToUpper() + @"\";
                }
                int backupday = 0;
                Int32.TryParse(Request.Form["backupday"].ToString(), out backupday);
                string cellcolor = Request.Form["cellcolor"];
                Pool[] pools = db.Pools.ToArray();
                foreach (Pool already in pools)
                {
                    if (already.Name.ToLower() == poolname.ToLower())
                    {
                        return "Ce nom de Pool existe déjà dans la base de données.";
                    }
                }
                Pool pool = db.Pools.Create();
                pool.Name = poolname;
                pool.BackupDay = backupday;
                if (backupday == 5 || backupday == 6) //If day is in weekend
                {
                    pool.CheckDay = 1; //The check must be executed on Mondays
                }
                else
                {
                    pool.CheckDay = backupday + 1; //Else it must be executed the next day
                }
                if (backupday == 7) //If day is everyday
                {
                    pool.CheckDay = 7; //The check must be executed Everydays
                }
                if (cellcolor != null)
                {
                    pool.CellColor = "#" + cellcolor;
                }
                else
                {
                    pool.CellColor = "#22ff22";
                }
                pool.BackupManager = backupmanager.ToUpper();
                pool.CheckFolder = checkfolder.ToUpper();
                if (ModelState.IsValid)
                {
                    db.Pools.Add(pool);
                    db.SaveChanges();
                    Specific_Logging(new Exception("...."), "AddPool " + pool.Name, 3);
                    return "Ce Pool a été rajouté à la liste.\n" +
                        "Nous vous invitons à remplir la liste des serveurs qu'il contient.";
                }
                Specific_Logging(new Exception("...."), "AddPool " + pool.Name, 2);
                return "Erreur lors de l'ajout du Pool";
            }
            catch (Exception exception)
            {
                Specific_Logging(exception, "AddPool");
                return "Erreur lors de l'ajout du Pool";
            }
        }

        public string EditPool(int id)
        {
            try
            {
                Pool pool = db.Pools.Find(id);
                string poolname = Request.Form["poolname"];
                string backupmanager = Request.Form["backupmanager"];
                string checkfolder = Request.Form["checkfolder"];
                if (backupmanager == null || backupmanager.Trim() == "")
                {
                    backupmanager = HomeController.BACKUP_REMOTE_SERVER_EXEC;
                }
                if (checkfolder == null || checkfolder.Trim() == "")
                {
                    checkfolder = HomeController.BACKUP_REMOTE_CHECK_SERVER_ROOT_FOLDER + poolname.ToUpper() + @"\";
                }
                int backupday = 0;
                Int32.TryParse(Request.Form["backupday"].ToString(), out backupday);
                string cellcolor = Request.Form["cellcolor"];
                Pool[] pools = db.Pools.ToArray();
                foreach (Pool already in pools)
                {
                    if ((already.Name.ToLower() == poolname.ToLower()) &&
                         already.Id != pool.Id)
                    {
                        return "Ce nom de Pool existe déjà dans la base de données.";
                    }
                }
                pool.Name = poolname;
                pool.BackupDay = backupday;
                if (backupday == 5 || backupday == 6) //If day is in weekend
                {
                    pool.CheckDay = 1; //The check must be executed on Mondays
                }
                else
                {
                    pool.CheckDay = backupday + 1; //Else it must be executed the next day
                }
                if (backupday == 7) //If day is everyday
                {
                    pool.CheckDay = 7; //The check must be executed Everydays
                }
                if (cellcolor != null)
                {
                    pool.CellColor = "#" + cellcolor;
                }
                pool.BackupManager = backupmanager.ToUpper();
                pool.CheckFolder = checkfolder.ToUpper();
                if (ModelState.IsValid)
                {
                    db.Entry(pool).State = System.Data.Entity.EntityState.Modified;
                    db.SaveChanges();
                    Specific_Logging(new Exception("...."), "EditPool " + pool.Name, 3);
                    return "Les modifications ont été effectuées sur le Pool.\n";
                }
                Specific_Logging(new Exception("...."), "EditPool " + pool.Name, 2);
                return "Erreur lors de la modification du Pool";
            }
            catch (Exception exception)
            {
                Specific_Logging(exception, "EditPool");
                return "Erreur lors de la modification du Pool";
            }
        }

        public string EditPoolAccounts(int id)
        {
            try
            {
                string author = User.Identity.Name;
                Pool pool = db.Pools.Find(id);
                string execution_account = Request.Form["execution_account"];
                string check_account = Request.Form["check_account"];
                string session_password = Request.Form["session_password"];
                if (!McoUtilities.IsValidLoginPassword(author, McoUtilities.Encrypt(session_password)))
                {
                    return "Authentification échouée:\n mauvaise combinaison username/password";
                }
                pool.CheckAccount = check_account.ToUpper();
                pool.ExecutionAccount = execution_account.ToUpper();
                string logs = "Check with " + pool.CheckAccount;
                logs += "Exec with " + pool.ExecutionAccount;
                if (check_account != null && check_account.Trim() != ""
                    && execution_account != null && execution_account.Trim() != "")
                {
                    if (ModelState.IsValid)
                    {
                        db.Entry(pool).State = System.Data.Entity.EntityState.Modified;
                        db.SaveChanges();
                        Specific_Logging(new Exception("...."), "EditPoolAccounts " + pool.Name + " " + logs, 2);
                        return "Les modifications ont été effectuées sur le Pool.\n";
                    }
                }
                else
                {
                    Specific_Logging(new Exception("...."), "EditPoolAccounts " + pool.Name, 2);
                }
                return "Erreur lors de la modification des comptes du Pool";
            }
            catch (Exception exception)
            {
                Specific_Logging(exception, "EditPoolAccounts");
                return "Erreur lors de la modification des comptes du Pool";
            }
        }

        public string GetAccountsList(int id)
        {
            Pool pool = db.Pools.Find(id);
            if (pool == null)
            {
                return Accounts_Controller.GetStringedAccountsList();
            }
            string options = "";
            List<Account> accounts = Accounts_Controller.GetAccountsList();
            foreach (Account account in accounts)
            {
                options += "<option val='" + account.DisplayName.ToUpper() + "'";
                if (pool.CheckAccount != null && account.DisplayName.ToUpper() == pool.CheckAccount.ToUpper())
                {
                    options += " selected";
                }
                options += ">" + account.DisplayName.ToUpper() + "</option>";
            }
            return options;
        }

        public string DeletePool(int id)
        {
            Pool pool = db.Pools.Find(id);
            BackupServer[] servers = pool.BackupServers.ToArray();
            foreach (BackupServer server in servers)
            {
                BackupServer_Report[] serverReports = server.BackupServer_Reports.ToArray();
                foreach (BackupServer_Report serverReport in serverReports)
                {
                    db.BackupServerReports.Remove(serverReport);
                }
                db.BackupServers.Remove(server);
            }
            db.Pools.Remove(pool);
            db.SaveChanges();
            Specific_Logging(new Exception("...."), "DeletePool " + pool.Name, 3);
            return "Le pool " + pool.Name + " ainsi que les serveurs qu'il contient ont été correctement supprimés";
        }

        [HttpPost]
        public string Import(bool import)
        {
            string message = "";
            if (import)
            {
                BackupReport[] reports = db.BackupReports.ToArray();
                foreach (BackupReport report in reports)
                {
                    message += DeleteBackupReport(report.Id) + "\n <br />";
                }

                Pool[] pools = db.Pools.ToArray();
                foreach (Pool pool in pools)
                {
                    message += DeletePool(pool.Id) + "\n <br />";
                }


                //Feed the database With The Pools
                MyApplication = new Excel.Application();
                MyApplication.Visible = false;

                string[] weekday = { "Dimanche", "Lundi", "Mardi", "Mercredi", "Jeudi", "Vendredi", "Samedi", "Quotidien" };
                try
                {
                    MyWorkbook = MyApplication.Workbooks.Open(HomeController.BESR_INIT_FILE);
                    MySheet = (Excel.Worksheet)MyWorkbook.Sheets["Pool"]; // Explicit cast is not required here
                    int lastRow = MySheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row;
                    Dictionary<int, ServersController.VirtualizedServer> FOREST = ServersController.GetInformationsFromForestDomains();
                    ReftechServers[] REFTECH_SERVERS = null;
                    try
                    {
                        REFTECH_SERVERS = db.ReftechServers.ToArray();
                    }
                    catch { }
                    for (int index = 2; index < lastRow; index++)
                    {
                        System.Array MyValues = (System.Array)MySheet.get_Range("A" +
                            index.ToString(), "H" + index.ToString()).Cells.Value;
                        Pool pool = db.Pools.Create();
                        if (MyValues.GetValue(1, 1) == null)
                        {
                            break;
                        }
                        pool.Name = MyValues.GetValue(1, 1).ToString();
                        int dayindex = 0;
                        foreach (string day in weekday)
                        {
                            if (day.ToLower() == MyValues.GetValue(1, 2).ToString().ToLower())
                            {
                                pool.BackupDay = dayindex;
                                break;
                            }
                            dayindex++;
                        }
                        if (pool.BackupDay == 6 || pool.BackupDay == 5)
                        {
                            pool.CheckDay = 1;
                        }
                        else
                        {
                            pool.CheckDay = pool.BackupDay + 1;
                        }
                        if (pool.BackupDay == 7)
                        {
                            pool.CheckDay = 7;
                        }
                        pool.CellColor = "#" + MyValues.GetValue(1, 3).ToString();
                        string backup_manager = (MyValues.GetValue(1, 4) != null) ? MyValues.GetValue(1, 4).ToString().ToUpper() :
                            HomeController.BACKUP_REMOTE_SERVER_EXEC;
                        string check_folder = (MyValues.GetValue(1, 5) != null) ? MyValues.GetValue(1, 5).ToString().ToUpper() :
                            HomeController.BACKUP_REMOTE_CHECK_SERVER_ROOT_FOLDER + pool.Name + @"\";
                        string execution_account = (MyValues.GetValue(1, 6) != null) ? MyValues.GetValue(1, 6).ToString() :
                            HomeController.DEFAULT_DOMAIN_AND_USERNAME_IMPERSONNATION;
                        string check_account = (MyValues.GetValue(1, 7) != null) ? MyValues.GetValue(1, 7).ToString() :
                            HomeController.DEFAULT_DOMAIN_AND_USERNAME_IMPERSONNATION;


                        pool.CheckFolder = check_folder.ToUpper();
                        pool.BackupManager = backup_manager.ToUpper();

                        pool.CheckAccount = check_account.ToUpper();
                        pool.ExecutionAccount = execution_account.ToUpper();
                        db.Pools.Add(pool);
                        db.SaveChanges();
                        message += "Le Pool " + pool.Name + " a été généré. \n <br />";
                    }

                    //Feed The database with servers              
                    MySheet = (Excel.Worksheet)MyWorkbook.Sheets["Servers"]; // Explicit cast is not required here
                    lastRow = MySheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row;
                    pools = db.Pools.ToArray();

                    for (int index = 2; index <= lastRow; index++)
                    {
                        System.Array MyValues = (System.Array)MySheet.get_Range("A" +
                            index.ToString(), "D" + index.ToString()).Cells.Value;

                        BackupServer server = db.BackupServers.Create();

                        if (MyValues.GetValue(1, 1) == null)
                        {
                            break;
                        }

                        server.Name = MyValues.GetValue(1, 2).ToString();
                        ServersController.VirtualizedServer_Result virtual_server = ServersController.GetServerInformations(FOREST, REFTECH_SERVERS, server, HomeController.BESR_MODULE, false);
                        server = virtual_server.BESR_Server;
                        server.Disks = MyValues.GetValue(1, 3).ToString();
                        string poolName = MyValues.GetValue(1, 1).ToString();
                        foreach (Pool pool in pools)
                        {
                            if (pool.Name.Trim() == poolName.Trim())
                            {
                                server.Pool = pool;
                                break;
                            }
                        }
                        db.BackupServers.Add(server);
                        db.SaveChanges();
                    }
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

        public string Export()
        {
            string message = "", FileName = "";
            //Feed the database With The Apps
            MyApplication = new Excel.Application();
            MyApplication.Visible = false;
            try
            {
                List<Pool> pools = db.Pools.OrderBy(name => name.Name).ToList();
                if (pools.Count != 0)
                {
                    //Feed the database With The Apps
                    MyApplication = new Excel.Application();
                    MyApplication.Visible = false;

                    MyWorkbook = MyApplication.Workbooks.Open(HomeController.BESR_DEFAULT_INIT_FILE_README);
                    MySheet = (Excel.Worksheet)MyWorkbook.Sheets.Add();
                    MySheet.Name = "Pool";
                    MySheet.Activate();

                    MySheet = (Excel.Worksheet)MyWorkbook.Sheets["Pool"];
                    int lastRow = MySheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row;
                    Excel.Range ActualRange = MySheet.get_Range("A" + lastRow,
                            "Z" + lastRow);

                    MySheet.Cells[lastRow, 1] = HomeController.BESR_MODULE;
                    MySheet.Cells[lastRow, 1].EntireColumn.Font.Bold = true;
                    MySheet.Cells[lastRow, 2] = "SaveDay";
                    MySheet.Cells[lastRow, 2].EntireColumn.Font.Bold = true;
                    MySheet.Cells[lastRow, 3] = "CellColor";
                    MySheet.Cells[lastRow, 3].EntireColumn.Font.Bold = true;
                    MySheet.Cells[lastRow, 4] = "Gestionnaire de sauvegarde";
                    MySheet.Cells[lastRow, 4].EntireColumn.Font.Bold = true;
                    MySheet.Cells[lastRow, 5] = "Répertoire de vérification";
                    MySheet.Cells[lastRow, 5].EntireColumn.Font.Bold = true;
                    MySheet.Cells[lastRow, 6] = "Compte d'exécution de sauvegarde";
                    MySheet.Cells[lastRow, 6].EntireColumn.Font.Bold = true;
                    MySheet.Cells[lastRow, 7] = "Compte d'accès au répertoire de vérification";
                    MySheet.Cells[lastRow, 7].EntireColumn.Font.Bold = true;

                    MySheet.Cells[lastRow, 1].EntireRow.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                    MySheet.Cells[lastRow, 1].EntireColumn.ColumnWidth = 25;
                    MySheet.Cells[lastRow, 1].Interior.Color = System.Drawing.ColorTranslator.FromHtml("#a3a3a3");
                    MySheet.Cells[lastRow, 2].EntireColumn.ColumnWidth = 15;
                    MySheet.Cells[lastRow, 2].Interior.Color = System.Drawing.ColorTranslator.FromHtml("#a3a3a3");
                    MySheet.Cells[lastRow, 3].EntireColumn.ColumnWidth = 35;
                    MySheet.Cells[lastRow, 3].Interior.Color = System.Drawing.ColorTranslator.FromHtml("#a3a3a3");
                    MySheet.Cells[lastRow, 4].EntireColumn.ColumnWidth = 35;
                    MySheet.Cells[lastRow, 5].EntireColumn.ColumnWidth = 35;
                    MySheet.Cells[lastRow, 6].EntireColumn.ColumnWidth = 35;
                    MySheet.Cells[lastRow, 7].EntireColumn.ColumnWidth = 35;

                    MySheet.Cells[lastRow, 1].EntireRow.Interior.Color = System.Drawing.ColorTranslator.FromHtml("#fff");
                    lastRow++;

                    string[] weekday = { "Dimanche", "Lundi", "Mardi", "Mercredi", "Jeudi", "Vendredi", "Samedi", "Quotidien" };
                    foreach (Pool pool in pools)
                    {
                        MySheet.Cells[lastRow, 1].EntireRow.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                        MySheet.Cells[lastRow, 1].Interior.Color = System.Drawing.ColorTranslator.FromHtml(pool.CellColor);
                        MySheet.Cells[lastRow, 1] = pool.Name;
                        MySheet.Cells[lastRow, 2] = weekday[pool.BackupDay];
                        MySheet.Cells[lastRow, 3] = pool.CellColor.Substring(1);
                        MySheet.Cells[lastRow, 4] = pool.BackupManager;
                        MySheet.Cells[lastRow, 5] = pool.CheckFolder;
                        MySheet.Cells[lastRow, 6] = pool.ExecutionAccount;
                        MySheet.Cells[lastRow, 7] = pool.CheckAccount;
                        lastRow++;
                    }

                    List<BackupServer> servers = db.BackupServers.OrderBy(name => name.Pool.Name).ToList();

                    MySheet = (Excel.Worksheet)MyWorkbook.Sheets.Add();
                    MySheet.Name = "Servers";
                    MySheet.Activate();

                    MySheet = (Excel.Worksheet)MyWorkbook.Sheets["Servers"];
                    lastRow = MySheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row;
                    ActualRange = MySheet.get_Range("A" + lastRow, "Z" + lastRow);
                    MySheet.Cells[lastRow, 1].EntireRow.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                    MySheet.Cells[lastRow, 1].EntireRow.Interior.Color = System.Drawing.ColorTranslator.FromHtml("#fff");
                    MySheet.Cells[lastRow, 1] = "POOL";
                    MySheet.Cells[lastRow, 1].EntireColumn.Font.Bold = true;
                    MySheet.Cells[lastRow, 1].EntireColumn.ColumnWidth = 25;
                    MySheet.Cells[lastRow, 2] = "SERVEUR";
                    MySheet.Cells[lastRow, 2].EntireColumn.Font.Bold = true;
                    MySheet.Cells[lastRow, 3].EntireColumn.ColumnWidth = 25;
                    MySheet.Cells[lastRow, 3] = "Disques";
                    MySheet.Cells[lastRow, 3].EntireColumn.ColumnWidth = 25;
                    MySheet.Cells[lastRow, 3].EntireColumn.Font.Bold = true;

                    lastRow++;

                    foreach (BackupServer server in servers)
                    {
                        MySheet.Cells[lastRow, 1].EntireRow.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                        MySheet.Cells[lastRow, 1] = server.Pool.Name;
                        MySheet.Cells[lastRow, 1].Interior.Color = System.Drawing.ColorTranslator.FromHtml(server.Pool.CellColor);
                        MySheet.Cells[lastRow, 2] = server.Name;
                        MySheet.Cells[lastRow, 2].Interior.Color = System.Drawing.ColorTranslator.FromHtml(server.Pool.CellColor);
                        MySheet.Cells[lastRow, 3] = server.Disks;
                        MySheet.Cells[lastRow, 3].Interior.Color = System.Drawing.ColorTranslator.FromHtml(server.Pool.CellColor);
                        lastRow++;
                    }

                    FileName = HomeController.BESR_RESULTS_FOLDER + "ExportBesr" + DateTime.Now.ToString("dd") +
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
                    message = "Base de données vide";
                }
            }
            catch (Exception exception)
            {
                Specific_Logging(exception, "Export");
                return "Erreur lors de l'ajout du Pool";
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

        public string DownloadInitFile()
        {
            try
            {
                System.Web.HttpResponse response = System.Web.HttpContext.Current.Response;
                response.ClearContent();
                response.Clear();
                response.ContentType = "text/plain";
                string filepath = HomeController.BESR_RELATIVE_INIT_FILE;
                response.AddHeader("Content-Disposition", "attachment; filename=" + filepath + ";");
                String RelativePath = HomeController.BESR_DEFAULT_INIT_FILE.Replace(Request.ServerVariables["APPL_PHYSICAL_PATH"], String.Empty);
                response.TransmitFile(HomeController.BESR_DEFAULT_INIT_FILE);
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

        [HttpPost]
        public string AddBackupServer(int id)
        {
            Pool pool = db.Pools.Find(id);
            if (pool == null)
            {
                return HttpNotFound().ToString();
            }
            try
            {
                string servername = Request.Form["servername"];
                string serverdisks = Request.Form["serverdisks"];

                serverdisks = serverdisks.Substring(0, serverdisks.Length - 2);
                BackupServer[] servers = db.BackupServers.ToArray();
                foreach (BackupServer already in servers)
                {
                    if (already.Name.ToLower().Trim() == servername.ToLower().Trim())
                    {
                        return "Ce serveur existe déjà dans le pool " + already.Pool.Name + "\n" +
                            "Veuillez d'arbord le supprimer de ce Pool avant de l'ajouter dans celui-ci.";
                    }
                }

                BackupServer server = db.BackupServers.Create();
                server.Name = servername;
                Dictionary<int, ServersController.VirtualizedServer> FOREST = ServersController.GetInformationsFromForestDomains();
                ReftechServers[] REFTECH_SERVERS = null;
                try
                {
                    REFTECH_SERVERS = db.ReftechServers.ToArray();
                }
                catch { }
                ServersController.VirtualizedServer_Result virtual_server = ServersController.GetServerInformations(FOREST, REFTECH_SERVERS, server, HomeController.BESR_MODULE, false);
                server = virtual_server.BESR_Server;
                server.Disks = serverdisks;
                server.Pool = pool;
                server.PoolId = pool.Id;
                if (ModelState.IsValid)
                {
                    db.BackupServers.Add(server);
                    db.SaveChanges();
                    UpdateBackupServerFile(server, "server");
                    Specific_Logging(new Exception("...."), "AddBackupServer " + server.Pool.Name + " " + server.Name, 3);
                    return "Le serveur a correctement été rajouté au Pool.\n";
                }
                Specific_Logging(new Exception("...."), "AddBackupServer " + server.Pool.Name + " " + server.Name, 2);
                return "Erreur lors de l'ajout du serveur";
            }
            catch (Exception exception)
            {
                Specific_Logging(exception, "AddBackupServer");
                return "Erreur lors de l'ajout du serveur";
            }
        }

        [HttpPost]
        public string EditBackupServer(int id)
        {
            try
            {
                BackupServer server = db.BackupServers.Find(id);
                string servername = Request.Form["servername"];
                string serverdisks = Request.Form["serverdisks"];
                serverdisks = serverdisks.Substring(0, serverdisks.Length - 2);

                BackupServer[] servers = db.BackupServers.ToArray();
                foreach (BackupServer already in servers)
                {
                    if ((already.Name.ToLower().Trim() == servername.ToLower().Trim()) &&
                        already.Id != server.Id)
                    {
                        return "Ce serveur existe déjà dans le pool " + already.Pool.Name + "\n" +
                            "Veuillez d'arbord le supprimer de ce Pool avant de renommer celui-ci.";
                    }
                }

                server.Name = servername;
                server.Disks = serverdisks;
                if (ModelState.IsValid)
                {
                    db.Entry(server).State = System.Data.Entity.EntityState.Modified;
                    db.SaveChanges();
                    Specific_Logging(new Exception("...."), "EditBackupServer " + server.Pool.Name + " " + server.Name, 3);
                    return "Les modifications ont été effectuées.\n";
                }
                Specific_Logging(new Exception("...."), "EditBackupServer " + server.Pool.Name + " " + server.Name, 2);
                return "Erreur lors de la modification du serveur";
            }
            catch (Exception exception)
            {
                Specific_Logging(exception, "EditBackupServer");
                return "Erreur lors de la modification du serveur";
            }
        }

        public string DeleteBackupServer(int id)
        {
            BackupServer server = db.BackupServers.Find(id);
            string log = server.Pool.Name + " " + server.Name;
            BackupServer_Report[] serverReports = server.BackupServer_Reports.ToArray();
            foreach (BackupServer_Report serverReport in serverReports)
            {
                db.BackupServerReports.Remove(serverReport);
            }
            db.BackupServers.Remove(server);
            db.SaveChanges();
            if (db.BackupReports.Count() > 0)
            {
                try
                {
                    BackupReport report = db.BackupReports.OrderByDescending(rep => rep.Id).First();
                    MyApplication = new Excel.Application();
                    MyApplication.Visible = false;
                    string fileName = report.ResultPath;
                    if (System.IO.File.Exists(fileName))
                    {
                        MyWorkbook = MyApplication.Workbooks.Open(fileName);
                        CalendarWeekRule weekRule = CalendarWeekRule.FirstDay;
                        DayOfWeek firstWeekDay = DayOfWeek.Monday;
                        Calendar calendar = System.Threading.Thread.CurrentThread.CurrentCulture.Calendar;
                        int currentWeek = calendar.GetWeekOfYear(DateTime.Now, weekRule, firstWeekDay);

                        string sheetName = "Semaine " + currentWeek.ToString();
                        bool foundedSheet = false;
                        foreach (Excel.Worksheet sheet in MyWorkbook.Sheets)
                        {
                            if (sheet.Name.StartsWith(sheetName))
                            {
                                sheetName = sheet.Name;
                                foundedSheet = true;
                                break;
                            }
                        }
                        if (foundedSheet)
                        {
                            MySheet = (Excel.Worksheet)MyWorkbook.Sheets[sheetName];
                            MySheet.Activate();
                            Excel.Range range = MySheet.get_Range("B1",
                                "B" + MySheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row);
                            foreach (Excel.Range cell in range)
                            {
                                if (cell.Value == server.Name)
                                {
                                    MySheet.Cells[cell.Row, 5] = "Serveur supprimé par l'utilisateur " + User.Identity.Name;
                                    MySheet.Cells[cell.Row, 6].EntireRow.Interior.Color = System.Drawing.ColorTranslator.FromHtml("#aca8a4");
                                    MySheet.Cells[cell.Row, 6].EntireRow.Font.Color = System.Drawing.ColorTranslator.FromHtml(ColorTranslator.ToHtml(ContrastColor("#aca8a4")));
                                    MySheet.Cells[cell.Row, 6] = "";
                                    MySheet.Cells[cell.Row, 7] = "";
                                    MySheet.Cells[cell.Row, 8] = "";
                                    MySheet.Cells[cell.Row, 9] = "";
                                    break;
                                }
                            }
                            MyWorkbook.Save();
                        }
                    }
                }
                catch { }
                finally
                {
                    McoUtilities.CloseExcel(MyApplication, MyWorkbook, MySheet);
                }
            }
            Specific_Logging(new Exception("...."), "DeleteBackupServer " + server.Pool.Name + " " + server.Name, 3);
            return "Le serveur " + server.Name + " a été supprimé";
        }

        public JsonResult OpenPoolEmail(int id)
        {
            Pool pool = db.Pools.Find(id);
            int backupemailId = 0;
            try
            {
                Int32.TryParse(Request.Form["backupemailId"].ToString(), out backupemailId);
            }
            catch { }
            Email email = db.Emails.Find(backupemailId);
            if (email == null || pool == null)
            {
                return Json(HttpNotFound(), JsonRequestBehavior.AllowGet);
            }
            string body = "<div><br /><span>Rapport de check pour le pool " + pool.Name + ".</span><br />" +
                "Ci dessous, un tableau récapitulatif de l'etat des sauvegardes Windows pour l'ensemble des serveurs testés. " +
                "NB: Le rapport global de la semaine en cours sera envoyé par mail.<br /><br />";
            body += "-------------------------------------------------------------------<br /><br />" +
                "Nombre de serveurs vérifiés : " + pool.BackupServers.Count + "<br />" +
                "-------------------------------------------------------------------<br /><br />";

            body += "<table style='position:relative;width:100%;'><thead>" +
                "<tr style='position:relative;width:100%;height:35px;background-color:#ecebeb'><th>Pools</th><th>Serveurs</th><th>Partitions</th><th>Statut</th>" +
                "<th>Erreurs</th><th>Ping</th><th>Sauvegardes</th></thead><tbody>";
            BackupServer_Report[] backupserverreports = db.BackupServerReports.Where(
                serverreportid => serverreportid.BackupReport.Id == email.Report.Id)
                .Where(serverreportid => serverreportid.BackupServer.PoolId == id).ToArray();
            foreach (BackupServer_Report serverreport in backupserverreports)
            {
                string okayColor = "", pingColor = "";
                Color font = ContrastColor(serverreport.BackupServer.Pool.CellColor);
                if (serverreport.State == "OK")
                {
                    okayColor = "#00b050";
                }
                else
                {
                    okayColor = "#ff0000";
                }
                string ping = (serverreport.Ping == "Ping OK" || serverreport.Ping == "Ping Success") ? "Ping OK" : "Ping KO";
                if (ping == "Ping OK")
                {
                    pingColor = "";
                }
                else
                {
                    pingColor = "background-color:#fff;color:#ff0000";
                }
                body += "<tr style='position:relative;text-align:center;font-weight:bold;color:" + ColorTranslator.ToHtml(font) + ";width:100%;min-height:40px;border:1px solid #000;background-color:" + serverreport.BackupServer.Pool.CellColor + ";'>" +
                        "<td style='position:relative;text-align:center;width:100px;border:5px solid " + serverreport.BackupServer.Pool.CellColor + ";'>" + serverreport.BackupServer.Pool.Name + "</td>" +
                        "<td style='position:relative;text-align:center;width:100px;border:5px solid " + serverreport.BackupServer.Pool.CellColor + ";'>" + serverreport.BackupServer.Name + "</td>" +
                        "<td style='position:relative;text-align:center;width:80px;border:5px solid " + serverreport.BackupServer.Pool.CellColor + ";'>" + serverreport.BackupServer.Disks + "</td>" +
                        "<td style='position:relative;text-align:center;width:40px;border:5px solid " + serverreport.BackupServer.Pool.CellColor + ";background-color:#fff;color:" + okayColor + "'>" + serverreport.State + "</td>";
                body += "<td style='position:relative;text-align:center;width:80px;border:5px solid " + serverreport.BackupServer.Pool.CellColor + ";'>" + serverreport.Details + "</td>";
                body += "<td style='position:relative;text-align:center;width:60px;border:5px solid " + serverreport.BackupServer.Pool.CellColor + ";" + pingColor + "'>" + ping + "</td>";
                body += "<td style='position:relative;text-align:center;width:100px;border:5px solid " + serverreport.BackupServer.Pool.CellColor + ";'>" + serverreport.Relaunched + "</td></tr>";
            }
            body += "</tbody></table>";
            Dictionary<string, string> result = new Dictionary<string, string>();
            result.Add("subject", "Check du Pool " + pool.Name);
            result.Add("recipient", email.Recipients);
            result.Add("body", body);

            return Json(result, JsonRequestBehavior.AllowGet);
        }

        public JsonResult OpenPoolsEmail(int id)
        {
            Email email = db.Emails.Find(id);
            if (email == null)
            {
                return Json(HttpNotFound(), JsonRequestBehavior.AllowGet);
            }
            Dictionary<string, string> result = new Dictionary<string, string>();
            result.Add("subject", email.Subject);
            result.Add("recipient", email.Recipients);
            result.Add("body", "");
            result.Add("failedservers", "");
            bool autobesr = false;
            string[] poolList;
            List<Pool> SelectedPools = new List<Pool>();

            List<BackupServer_Report> FailedServers = new List<BackupServer_Report>();
            try
            {
                bool.TryParse(Request.Form["autobesr"], out autobesr);
                poolList = Request.Form["pools"].ToString().Split(';');
                foreach (string poolId in poolList)
                {
                    int ids = 0;
                    Int32.TryParse(poolId.Split('-')[1], out ids);
                    Pool pool = db.Pools.Find(ids);
                    if (pool != null)
                    {
                        SelectedPools.Add(pool);
                    }
                }
            }
            catch (Exception exception)
            {
                return Json(exception.Message, JsonRequestBehavior.AllowGet);
            }

            foreach (Pool pool in SelectedPools)
            {
                string body = "<div><br /><span>Check du Pool " + pool.Name + ".</span><br />" +
                "NB: Le rapport global de la semaine en cours sera envoyé par mail.<br />";
                body += "-------------------------------------------------------------------<br /><br />" +
                    "Nombre de serveurs vérifiés dans le Pool : " + pool.BackupServers.Count + "<br />" +
                    "Nombre de serveurs en erreur : " + db.BackupServerReports.Where(
                    serverreportid => serverreportid.BackupReport.Id == email.Report.Id)
                    .Where(serverreportid => serverreportid.BackupServer.PoolId == pool.Id).Where(serverstate => serverstate.State != "OK").Count().ToString() + "<br />" +
                    "-------------------------------------------------------------------<br />";

                body += "<table style='position:relative;width:100%;'><thead>" +
                    "<tr style='position:relative;width:100%;height:35px;background-color:#ecebeb'><th>Pools</th><th>Serveurs</th><th>Partitions</th><th>Statut</th>" +
                    "<th>Erreurs</th><th>Ping</th><th>Sauvegardes</th></thead><tbody>";
                BackupServer_Report[] backupserverreports = db.BackupServerReports.Where(
                    serverreportid => serverreportid.BackupReport.Id == email.Report.Id)
                    .Where(serverreportid => serverreportid.BackupServer.PoolId == pool.Id).ToArray();
                foreach (BackupServer_Report serverreport in backupserverreports)
                {
                    if (serverreport.State != "OK")
                    {
                        FailedServers.Add(serverreport);
                    }
                    string okayColor = "", pingColor = "";
                    Color font = ContrastColor(serverreport.BackupServer.Pool.CellColor);
                    if (serverreport.State == "OK")
                    {
                        okayColor = "#00b050";
                    }
                    else
                    {
                        okayColor = "#ff0000";
                    }
                    string ping = (serverreport.Ping == "Ping OK" || serverreport.Ping == "Ping Success") ? "Ping OK" : "Ping KO";
                    if (ping == "Ping OK")
                    {
                        pingColor = "";
                    }
                    else
                    {
                        pingColor = "background-color:#fff;color:#ff0000";
                    }
                    body += "<tr style='position:relative;text-align:center;font-weight:bold;color:" + ColorTranslator.ToHtml(font) + ";width:100%;min-height:40px;border:1px solid #000;background-color:" + serverreport.BackupServer.Pool.CellColor + ";'>" +
                        "<td style='position:relative;text-align:center;width:100px;border:5px solid " + serverreport.BackupServer.Pool.CellColor + ";'>" + serverreport.BackupServer.Pool.Name + "</td>" +
                        "<td style='position:relative;text-align:center;width:100px;border:5px solid " + serverreport.BackupServer.Pool.CellColor + ";'>" + serverreport.BackupServer.Name + "</td>" +
                        "<td style='position:relative;text-align:center;width:80px;border:5px solid " + serverreport.BackupServer.Pool.CellColor + ";'>" + serverreport.BackupServer.Disks + "</td>" +
                        "<td style='position:relative;text-align:center;width:40px;border:5px solid " + serverreport.BackupServer.Pool.CellColor + ";background-color:#fff;color:" + okayColor + "'>" + serverreport.State + "</td>";
                    body += "<td style='position:relative;text-align:center;width:80px;border:5px solid " + serverreport.BackupServer.Pool.CellColor + ";'>" + serverreport.Details + "</td>";
                    body += "<td style='position:relative;text-align:center;width:60px;border:5px solid " + serverreport.BackupServer.Pool.CellColor + ";" + pingColor + "'>" + ping + "</td>";
                    body += "<td style='position:relative;text-align:center;width:100px;border:5px solid " + serverreport.BackupServer.Pool.CellColor + ";'>" + serverreport.Relaunched + "</td></tr>";
                }
                body += "</tbody></table>";
                result["body"] += body + "<br />";
            }

            if (!autobesr)
            {
                string failedservers = "<div id='autobesr-div'><div class='backup-launcher-div'><span>" +
                    "Vous pouvez relancer la sauvegarde sur plusieurs serveurs en cochant leur <span class='nav-info'>checkbox</span> (case à cocher)" +
                    "<br />et en appuyant le bouton <span class='nav-info'>lancer sauvegarde</span><br /></span>" +
                    "<input id='backup-launcher-btn' class='backup-launcher-btn' type='button' value='Lancer Sauvegarde' /></div>" +
                    "<div class='failed-backupserver-list'>" +
                    "<table id='failed-backupserver-table' class='item-summary failed-backupserver-table'>";
                failedservers += "<thead>" +
                    "<tr style='position:relative;width:100%;height:35px;background-color:#ff3f3f'><th> </th><th>Pools</th><th>Serveurs</th><th>Etat</th><th>Ping</th>" +
                    "<th>Services</th><th>Sauvegarde</th><th>Erreurs</th><th>Actions</th></thead><tbody>";
                foreach (BackupServer_Report serverreport in FailedServers)
                {
                    failedservers += "<tr style='position:relative;color:#000;width:100%;min-height:40px;border:1px solid #000;background-color:" + serverreport.BackupServer.Pool.CellColor + ";'>";
                    failedservers += "<td style='position:relative;width:40px;border:1px solid #000;'><input type='text' disabled = 'disabled' class='action-id-getter' value='" +
                        serverreport.BackupServer.Id + "' id='action-id-sgetter-" + serverreport.BackupServer.Id + "'/>";
                    failedservers += "<input type='checkbox' class='selected-servers' name= '" + serverreport.BackupServer.Name + "' /></td>";
                    failedservers += "<td style='position:relative;width:100px;border:1px solid #000;'>" + serverreport.BackupServer.Pool.Name + "</td>" +
                        "<td style='position:relative;border:1px solid #000;'>" + serverreport.BackupServer.Name + "</td>" +
                        "<td style='position:relative;border:1px solid #000;'>" + serverreport.State + "</td>";
                    failedservers += "<td style='position:relative;border:1px solid #000;'>" + serverreport.Ping + "</td>";
                    failedservers += "<td style='position:relative;border:1px solid #000;'>" + serverreport.Services + "</td>";
                    failedservers += "<td style='position:relative;border:1px solid #000;'>" + serverreport.Relaunched + "</td>";
                    failedservers += "<td style='position:relative;border:1px solid #000;'>" + serverreport.Details + "</td>";
                    failedservers += "<td style='margin=0px;padding=0px;'>" +
                        "<div class='action-div'>" +
                        "<Input type='text' id='action-id-getter-" + serverreport.BackupServer.Id + "'" +
                        "value='" + serverreport.BackupServer.Id + "' disabled = 'disabled' class='action-id-getter' />" +
                        "<div class='action-icon backup-launcher' title='Relancer la sauvegarde pour ce serveur'></div>" +
                        "<div class='action-icon deleter' title='supprimer le serveur " + serverreport.BackupServer.Name + "'></div>" +
                        "</div></td></tr>";
                }
                failedservers += "</tbody></table></div></div>";
                result["failedservers"] += failedservers + "<br />";
            }
            return Json(result, JsonRequestBehavior.AllowGet);
        }

        public JsonResult ExecuteSchedule(int id)
        {
            BackupSchedule schedule = db.BackupSchedules.Find(id);
            if (schedule == null)
            {
                Specific_Logging(new Exception("...."), "ExecuteSchedule");
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

            UpdatePoolDatabase();
            Dictionary<string, string> results = new Dictionary<string, string>();
            results["response"] = "";
            results["email"] = "";
            results["errors"] = "";
            results["pools"] = "";
            results["autobesr"] = "";

            DateTime now = DateTime.Now;
            //DateTime yesterday = DateTime.Now.AddDays(-1);
            DateTime yesterday = DateTime.Now.AddDays(-1);
            int dayOfWeek = (int)now.Date.DayOfWeek;
            IQueryable<Pool> pools = db.Pools.Where(ide => ide.CheckDay == dayOfWeek);
            List<Pool> SelectedPools = new List<Pool>();
            foreach (Pool pool in pools)
            {
                SelectedPools.Add(pool);
            }
            if (SelectedPools.Count == 0)
            {
                results["response"] = "Aucun Pool n'a été sélectionné dans la base de données";
                results["email"] = "";
                results["errors"] = "";
            }
            int emailId = 0;
            string ExecutionErrors = "";
            try
            {
                MyApplication = new Excel.Application();
                MyApplication.Visible = false;
                MyApplication.DisplayAlerts = false;
                bool foundedFile = false;
                string[] report_files = System.IO.Directory.GetFiles(HomeController.BESR_RESULTS_FOLDER,
                    "Controle_hebdomadaire_sauvegarde_BESR_PCI_" + DateTime.Now.ToString("yyyy") + "*.xlsx", System.IO.SearchOption.TopDirectoryOnly);
                string filename = "Controle_hebdomadaire_sauvegarde_BESR_PCI_" + DateTime.Now.ToString("yyyy") + ".xlsx";
                if (report_files.Length > 0)
                {
                    filename = Directory.GetFiles(HomeController.BESR_RESULTS_FOLDER,
                        "Controle_hebdomadaire_sauvegarde_BESR_PCI_" + DateTime.Now.ToString("yyyy") + "*.xlsx", System.IO.SearchOption.TopDirectoryOnly)
                            .Select(x => new FileInfo(x))
                            .OrderByDescending(x => x.LastWriteTime).FirstOrDefault().FullName;

                    MyWorkbook = MyApplication.Workbooks.Open(filename);
                    foundedFile = true;
                }
                else
                {
                    ExecutionErrors += HomeController.BESR_RESULTS_FOLDER + "Controle_hebdomadaire_sauvegarde_BESR_PCI_" + DateTime.Now.ToString("yyyy") + "-(id).xlsx n'a pas été trouvé.\r\n";
                    MyWorkbook = MyApplication.Workbooks.Add(Excel.XlWBATemplate.xlWBATWorksheet);
                    foundedFile = false;
                }

                CalendarWeekRule weekRule = CalendarWeekRule.FirstDay;
                DayOfWeek firstWeekDay = DayOfWeek.Monday;
                Calendar calendar = System.Threading.Thread.CurrentThread.CurrentCulture.Calendar;
                int currentWeek = calendar.GetWeekOfYear(DateTime.Now, weekRule, firstWeekDay);

                string sheetName = "Semaine " + currentWeek.ToString();
                bool foundedSheet = false;
                foreach (Excel.Worksheet sheet in MyWorkbook.Sheets)
                {
                    if (sheet.Name.StartsWith(sheetName))
                    {
                        sheetName = sheet.Name;
                        foundedSheet = true;
                        break;
                    }
                }
                if (foundedSheet)
                {
                    MySheet = (Excel.Worksheet)MyWorkbook.Sheets[sheetName];
                }
                else
                {
                    MySheet = (Excel.Worksheet)MyWorkbook.Sheets.Add();
                    MySheet.Name = sheetName + " " + DateTime.Now.ToString("dd") +
                        DateTime.Now.ToString("MM") + DateTime.Now.ToString("yyyy");
                }

                MySheet.Activate();

                //Database objects init
                IQueryable<BackupReport> oldReports = db.BackupReports.Where(reportid => reportid.WeekNumber == currentWeek); //.First();
                BackupReport report;
                Email email;

                if (oldReports.Count() == 0)
                {
                    report = db.BackupReports.Create();
                    report.DateTime = DateTime.Now;
                    report.LastUpdate = DateTime.Now;
                    report.WeekNumber = currentWeek;
                    report.TotalChecked = 0;
                    report.TotalErrors = 0;
                    report.ResultPath = "";
                    report.Author = User.Identity.Name;
                    report.Module = HomeController.BESR_MODULE;
                    report.ScheduleId = schedule.Id;
                    report.Schedule = schedule;
                    email = db.Emails.Create();
                    report.Email = email;
                    email.Module = HomeController.BESR_MODULE;
                    email.Report = report;
                    email.Recipients = "";
                    email = Emails_Controller.SetRecipients(email, HomeController.BESR_MODULE);
                    if (ModelState.IsValid)
                    {
                        db.BackupReports.Add(report);
                        db.SaveChanges();
                        emailId = report.Email.Id;
                        int reportNumber = db.BackupReports.Count();
                        if (reportNumber > HomeController.BESR_MAX_REPORT_NUMBER)
                        {
                            int reportNumberToDelete = reportNumber - HomeController.BESR_MAX_REPORT_NUMBER;
                            BackupReport[] reportsToDelete =
                                db.BackupReports.OrderBy(ide => ide.Id).Take(reportNumberToDelete).ToArray();
                            foreach (BackupReport toDeleteReport in reportsToDelete)
                            {
                                DeleteBackupReport(toDeleteReport.Id);
                            }
                        }
                    }
                    else
                    {
                        results["response"] = "KO";
                        results["email"] = null;
                        results["errors"] = "Impossible de créer un rapport dans la base de données.";
                        goto CHECK_FINALIZER;
                    }
                }
                else
                {
                    report = oldReports.First();
                    report.ScheduleId = schedule.Id;
                    report.Schedule = schedule;
                    report.DateTime = DateTime.Now;
                    email = report.Email;
                    email.Recipients = "";
                    email = Emails_Controller.SetRecipients(email, HomeController.BESR_MODULE);
                    if (ModelState.IsValid)
                    {
                        db.Entry(report).State = System.Data.Entity.EntityState.Modified;
                        db.SaveChanges();
                        emailId = report.Email.Id;
                    }
                    else
                    {
                        results["response"] = "KO";
                        results["email"] = null;
                        results["errors"] = "Impossible de créer un rapport dans la base de données.";
                        goto CHECK_FINALIZER;
                    }
                }

                //End database objects init
                int lastRow = MySheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row;
                ICollection<BackupServer> list = db.BackupServers.OrderBy(ide => ide.Pool.Id).ToArray();
                if (!foundedSheet)
                {
                    foreach (BackupServer server in list)
                    {
                        Color rangeColor = ContrastColor(server.Pool.CellColor);
                        Excel.Range ActualRange = MySheet.get_Range("A" + lastRow,
                                "M" + lastRow);
                        ActualRange.Interior.Color = System.Drawing.ColorTranslator.FromHtml(server.Pool.CellColor);
                        ActualRange.Font.Color = System.Drawing.ColorTranslator.FromHtml(ColorTranslator.ToHtml(rangeColor));
                        MySheet.Cells[lastRow, 1] = server.Pool.Name;
                        MySheet.Cells[lastRow, 1].EntireColumn.ColumnWidth = 15;
                        MySheet.Cells[lastRow, 2] = server.Name;
                        MySheet.Cells[lastRow, 2].EntireColumn.ColumnWidth = 15;
                        MySheet.Cells[lastRow, 3] = server.Disks;
                        MySheet.Cells[lastRow, 3].EntireColumn.ColumnWidth = 10;
                        MySheet.Cells[lastRow, 3].EntireRow.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                        MySheet.Cells[lastRow, 4].EntireColumn.ColumnWidth = 5;
                        lastRow += 1;
                    }
                    Excel.Range styling = MySheet.get_Range("E:E", System.Type.Missing);
                    styling.EntireColumn.ColumnWidth = 60;
                }

                foreach (Pool pool in SelectedPools)
                {
                    DayOfWeek[] week = new[] { DayOfWeek.Sunday, DayOfWeek.Monday, DayOfWeek.Tuesday,
                                     DayOfWeek.Wednesday,DayOfWeek.Thursday,DayOfWeek.Friday,
                                     DayOfWeek.Saturday};
                    if (pool.BackupDay == 7)
                    {
                        yesterday = DateTime.Now.AddDays(-1);
                    }
                    else
                    {
                        for (int day = -7; day < 0; day++)
                        {
                            yesterday = DateTime.Now.AddDays(day);
                            if (yesterday.DayOfWeek == week[pool.BackupDay])
                            {
                                break;
                            }
                        }
                    }

                    ICollection<BackupServer> servers = db.BackupServers.Where(ide => ide.Pool.Id == pool.Id).ToArray();

                    foreach (BackupServer server in servers)
                    {
                        int row_index = 5;
                        Excel.Range range = MySheet.get_Range("B1",
                            "Z" + MySheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row);

                        Dictionary<string, string> result = CheckBackupServer(server, yesterday, range);

                        BackupServer_Report serverReport = RegisterResult(report, server, result, false);
                        if (schedule.AutoRelaunch && serverReport.State != "OK")
                        {
                            if (serverReport.Ping != "Ping Success")
                            {
                                serverReport.Services = "Services: Ping KO";
                                serverReport.Relaunched = "Non relancées: Ping KO";
                            }
                            else
                            {
                                string services = ServiceLauncher(server.Id);
                                string BESROK = "Backup Exec System Recovery: Running: Manual";
                                string SYMOK = "SymSnapService: Running: Manual";
                                if ((services.IndexOf(BESROK, StringComparison.OrdinalIgnoreCase) > 0) &&
                                    services.IndexOf(SYMOK, StringComparison.OrdinalIgnoreCase) > 0)
                                {
                                    serverReport.Services = "OK";
                                    string relaunched = BackupExecLauncher(server.Id);
                                    if (relaunched.IndexOf("Sauvegarde OK") != -1)
                                    {
                                        serverReport.Relaunched = "Relancées";
                                    }
                                }
                                else
                                {
                                    serverReport.Services = "KO";
                                    serverReport.Relaunched = "Non relancées: Services non lancés.";
                                }
                            }
                            if (ModelState.IsValid)
                            {
                                db.Entry(serverReport).State = System.Data.Entity.EntityState.Modified;
                                db.SaveChanges();
                            }
                        }
                        UpdateCells(MySheet, server.Name, serverReport, row_index);
                    }
                }

            CHECK_FINALIZER:
                {
                    try
                    {
                        MyWorkbook.SaveAs(HomeController.BESR_RESULTS_FOLDER + "Controle_hebdomadaire_sauvegarde_BESR_PCI_" + DateTime.Now.ToString("yyyy") + "-" +
                            report.Id.ToString() + ".xlsx",
                            Excel.XlFileFormat.xlOpenXMLWorkbook, null, null, null, null,
                            Excel.XlSaveAsAccessMode.xlNoChange, null, null, null, null, null);

                        report.ResultPath = HomeController.BESR_RESULTS_FOLDER + "Controle_hebdomadaire_sauvegarde_BESR_PCI_" + DateTime.Now.ToString("yyyy") + "-" +
                            report.Id.ToString() + ".xlsx";
                    }
                    catch (Exception saveException)
                    {
                        Specific_Logging(saveException, "ExecuteSchedule");
                        ExecutionErrors += "Erreur de sauvegarde: " + saveException.Message + "\r\n";
                        try
                        {
                            MyWorkbook.SaveAs(HomeController.BESR_RESULTS_FOLDER + "Controle_hebdomadaire_sauvegarde_BESR_PCI_" + DateTime.Now.ToString("yyyy") + ".xlsx",
                                Excel.XlFileFormat.xlOpenXMLWorkbook, null, null, null, null,
                                Excel.XlSaveAsAccessMode.xlNoChange, null, null, null, null, null);
                            report.ResultPath = HomeController.BESR_RESULTS_FOLDER + "Controle_hebdomadaire_sauvegarde_BESR_PCI_" + DateTime.Now.ToString("yyyy") + ".xlsx";
                        }
                        catch { }
                    }
                    report.Duration = DateTime.Now.Subtract(report.DateTime);
                    report.LastUpdate = DateTime.Now;
                    report.TotalChecked = report.BackupServer_Reports.Count;
                    report.TotalErrors = report.BackupServer_Reports.Where(serverreport => serverreport.State != "OK").Count();
                    if (ModelState.IsValid)
                    {
                        db.Entry(report).State = System.Data.Entity.EntityState.Modified;
                        db.SaveChanges();
                        string buildOk = BuildBackupEmail(email.Id);
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
            }
            catch (Exception running)
            {
                Specific_Logging(running, "ExecuteSchedule");
            }
            finally
            {
                McoUtilities.CloseExcel(MyApplication, MyWorkbook, MySheet);
            }
            Emails_Controller.AutoSend(emailId);
            results["response"] = "OK";
            results["email"] = emailId.ToString();
            results["errors"] = "Fin d'exécution. \n" + "Erreurs d'exécution : " + ExecutionErrors;
            schedule.State = (schedule.Multiplicity != "Une fois") ? "Planifié" : "Terminé";
            schedule.NextExecution = Schedules_Controller.GetNextExecution(schedule);
            schedule.Executed++;
            if (ModelState.IsValid)
            {
                db.Entry(schedule).State = System.Data.Entity.EntityState.Modified;
                db.SaveChanges();
            }
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
            McoUtilities.Specific_Logging(exception, action, HomeController.BESR_MODULE, level, author);
        }
    }
}
