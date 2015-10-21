using System;
using System.Linq;
using System.Web.Mvc;
using Microsoft.Win32.TaskScheduler;

namespace McoEasyTool.Controllers
{
    [AllowAnonymous]
    public class SchedulesController : Controller
    {
        private DataModelContainer db = new DataModelContainer();

        public string ReSend(Schedule schedule)
        {
            if (schedule != null)
            {
                if (db.Reports.Where(report => report.ScheduleId == schedule.Id).Count() != 0)
                {
                    Report report = db.Reports.Where(rep => rep.ScheduleId == schedule.Id).OrderByDescending(rep => rep.Id).First();
                    ReportsController Reports_Controller = new ReportsController();
                    McoUtilities.General_Logging(new Exception("...."), "ReSend", 3);
                    return Reports_Controller.ReSend(report.Id);
                }
                McoUtilities.General_Logging(new Exception("...."), "ReSend No report", 3);
                return "Cette tâche planifiée n'a pour l'instant généré aucun rapport, ou alors ils ont été supprimés.";
            }
            McoUtilities.General_Logging(new Exception("...."), "ReSend Schedule", 2);
            return "Cette tâche planifiée n'a pas été trouvée dans la base de données.";
        }

        public string Create(Schedule schedule)
        {
            using (TaskService taskservice = new TaskService())
            {
                TaskDefinition taskdefinition = taskservice.NewTask();
                taskdefinition.RegistrationInfo.Description = "Lancement automatique ";
                if (schedule.Multiplicity == "Quotidien")
                {
                    taskdefinition.Triggers.Add(new WeeklyTrigger(DaysOfTheWeek.Monday | DaysOfTheWeek.Tuesday | DaysOfTheWeek.Wednesday |
                    DaysOfTheWeek.Thursday | DaysOfTheWeek.Friday) { StartBoundary = Convert.ToDateTime(schedule.NextExecution) });
                }
                else
                {
                    if (schedule.Multiplicity == "Hebdomadaire")
                    {
                        taskdefinition.Triggers.Add(new WeeklyTrigger { StartBoundary = Convert.ToDateTime(schedule.NextExecution) });
                    }
                    else
                    {
                        Trigger trigger = Trigger.CreateTrigger(TaskTriggerType.Time);
                        trigger.StartBoundary = Convert.ToDateTime(schedule.NextExecution);
                        taskdefinition.Triggers.Add(trigger);
                        taskdefinition.Settings.DeleteExpiredTaskAfter.Add(new TimeSpan(1, 0, 0));
                    }
                }
                taskdefinition.Principal.UserId = HomeController.DEFAULT_DOMAIN_AND_USERNAME_IMPERSONNATION;
                taskdefinition.Settings.WakeToRun = true;
                taskdefinition.Settings.StartWhenAvailable = true;
                taskdefinition.Settings.StopIfGoingOnBatteries = false;
                taskdefinition.Actions.Add(new ExecAction(HomeController.MCO_SCHEDULER_TOOL, schedule.Module + " CHECK " + schedule.Id.ToString()));
                taskservice.RootFolder.RegisterTaskDefinition(schedule.TaskName, taskdefinition,
                    TaskCreation.CreateOrUpdate, HomeController.DEFAULT_DOMAIN_AND_USERNAME_IMPERSONNATION,
                    McoUtilities.Decrypt(HomeController.DEFAULT_PASSWORD_IMPERSONNATION));
                McoUtilities.General_Logging(new Exception("...."), "Create " + schedule.Module + " " + schedule.TaskName, 3);
                return "Une analyse a été programmée à la date spécifiée";
            }
        }

        public string Edit(Schedule schedule)
        {
            using (TaskService taskservice = new TaskService())
            {
                TaskDefinition taskdefinition = taskservice.NewTask();
                taskdefinition.RegistrationInfo.Description = "Lancement automatique ";
                if (schedule.Multiplicity == "Quotidien")
                {
                    taskdefinition.Triggers.Add(new WeeklyTrigger(DaysOfTheWeek.Monday | DaysOfTheWeek.Tuesday | DaysOfTheWeek.Wednesday |
                    DaysOfTheWeek.Thursday | DaysOfTheWeek.Friday) { StartBoundary = Convert.ToDateTime(schedule.NextExecution) });
                }
                else
                {
                    if (schedule.Multiplicity == "Hebdomadaire")
                    {
                        taskdefinition.Triggers.Add(new WeeklyTrigger { StartBoundary = Convert.ToDateTime(schedule.NextExecution) });
                    }
                    else
                    {
                        Trigger trigger = Trigger.CreateTrigger(TaskTriggerType.Time);
                        trigger.StartBoundary = Convert.ToDateTime(schedule.NextExecution);
                        taskdefinition.Triggers.Add(trigger);
                        taskdefinition.Settings.DeleteExpiredTaskAfter.Add(new TimeSpan(1, 0, 0));
                    }

                }
                taskdefinition.Principal.UserId = HomeController.DEFAULT_DOMAIN_AND_USERNAME_IMPERSONNATION;
                taskdefinition.Settings.WakeToRun = true;
                taskdefinition.Settings.StartWhenAvailable = true;
                taskdefinition.Settings.StopIfGoingOnBatteries = false;
                taskdefinition.Actions.Add(new ExecAction(HomeController.MCO_SCHEDULER_TOOL, schedule.Module + " CHECK " + schedule.Id.ToString()));
                taskservice.RootFolder.RegisterTaskDefinition(schedule.TaskName, taskdefinition,
                    TaskCreation.CreateOrUpdate, HomeController.DEFAULT_DOMAIN_AND_USERNAME_IMPERSONNATION,
                    McoUtilities.Decrypt(HomeController.DEFAULT_PASSWORD_IMPERSONNATION));
                McoUtilities.General_Logging(new Exception("...."), "Edit " + schedule.Module + " " + schedule.TaskName, 3);
                return "La tâche a été correctement modifiée";
            }
        }

        public string Delete(Schedule schedule)
        {
            try
            {
                using (TaskService taskservice = new TaskService())
                {
                    string module = schedule.Module;
                    taskservice.RootFolder.DeleteTask(schedule.TaskName, false);
                    McoUtilities.General_Logging(new Exception("...."), "Delete " + module + " " + schedule.TaskName);
                }
            }
            catch (NullReferenceException) { }
            catch (Exception exception)
            {
                McoUtilities.General_Logging(exception, "Delete Schedule", 0);
                return "Une erreur est surveunue lors de la suppression\n" +
                    exception.Message;
            }
            McoUtilities.General_Logging(new Exception("...."), "Delete Schedule", 2);
            return "La tâche a été correctement supprimée";
        }

        public DateTime GetNextExecution(Schedule schedule)
        {
            DateTime next_execution = schedule.NextExecution.Value;
            DateTime now = DateTime.Now;
            if (now.CompareTo(next_execution) == -1)
            {
                return next_execution;
            }
            DateTime next =
                new DateTime(now.Year, now.Month, now.Day, next_execution.Hour, next_execution.Minute, next_execution.Second);
            switch (schedule.Multiplicity)
            {
                case "Quotidien":
                    next = next.AddDays(1);
                    if (next.DayOfWeek == DayOfWeek.Saturday)
                    {
                        next = next.AddDays(2);
                    }
                    if (next.DayOfWeek == DayOfWeek.Sunday)
                    {
                        next = next.AddDays(1);
                    }
                    break;
                case "Hebdomadaire":
                    next = next.AddDays(7);
                    break;
                default:
                    break;
            }
            return next;
        }

        protected override void Dispose(bool disposing)
        {
            db.Dispose();
            base.Dispose(disposing);
        }
    }
}