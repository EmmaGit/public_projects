using System;
using System.Collections.Generic;
using System.Data;
using System.Net.Mail;
using System.Web.Mvc;
using System.Linq;

namespace McoEasyTool.Controllers
{
    [AllowAnonymous]
    public class EmailsController : Controller
    {
        private DataModelContainer db = new DataModelContainer();

        public JsonResult Open(int id)
        {
            Email email = db.Emails.Find(id);
            if (email == null)
            {
                return Json(HttpNotFound(), JsonRequestBehavior.AllowGet);
            }
            Dictionary<string, string> result = new Dictionary<string, string>();
            result.Add("subject", email.Subject);
            result.Add("recipient", email.Recipients);
            result.Add("body", email.Body);
            return Json(result, JsonRequestBehavior.AllowGet);
        }

        public JsonResult Send(int id)
        {
            try
            {
                Email email = db.Emails.Find(id);
                if (email == null)
                {
                    return Json(HttpNotFound(), JsonRequestBehavior.AllowGet);
                }

                string recipients = Request.Form["Recipient"].ToString();
                string subject = Request.Form["Subject"].ToString();
                string message = Request.Form["Message"].ToString();

                email.Recipients = recipients;
                email.Subject = subject;
                email.Body = message + "<br /><br />" + email.Body;

                Dictionary<string, string> result = new Dictionary<string, string>();
                result.Add("Response", null);
                result.Add("Email", null);

                if (ModelState.IsValid)
                {
                    db.Entry(email).State = System.Data.Entity.EntityState.Modified;
                    db.SaveChanges();
                    Report report = db.Reports.Find(email.Report.Id);
                    TimeSpan Duration = DateTime.Now.Subtract(report.DateTime);
                    report.Duration = Duration;
                    if (ModelState.IsValid)
                    {
                        db.Entry(report).State = System.Data.Entity.EntityState.Modified;
                        db.SaveChanges();
                        result["Email"] = "Changement(s) sauvegardé(s)";
                    }
                }

                SmtpClient client = new SmtpClient();
                client.Port = HomeController.MAIL_PORT_NUMBER;
                client.DeliveryMethod = SmtpDeliveryMethod.Network;
                client.UseDefaultCredentials = false;
                client.Host = HomeController.MAIL_HOST_SERVER;


                MailMessage mail = new MailMessage();
                mail.From = new MailAddress(HomeController.MAIL_ADDRESS_SENDER);
                mail.IsBodyHtml = true;
                mail.Body = email.Body;
                mail.Attachments.Add(new Attachment(@email.Report.ResultPath));
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
                if (ModelState.IsValid)
                {
                    db.Entry(email).State = System.Data.Entity.EntityState.Modified;
                    db.SaveChanges();
                }
                result["Response"] = "Votre email a été correctement envoyé aux personnes sélectionnées.";
                McoUtilities.General_Logging(new Exception("...."), "Send Email", 3);
                return Json(result, JsonRequestBehavior.AllowGet);
            }
            catch (Exception exception)
            {
                McoUtilities.General_Logging(exception, "Send Email", 0);
                Dictionary<string, string> result = new Dictionary<string, string>();
                result.Add("Response", "Une erreur est survenue durant l'envoi du message");
                result.Add("Error", exception.Message);
                return Json(result, JsonRequestBehavior.AllowGet);
            }
        }

        public JsonResult ReSend(int id)
        {
            try
            {
                Email email = db.Emails.Find(id);
                if (email == null)
                {
                    return Json(HttpNotFound(), JsonRequestBehavior.AllowGet);
                }

                Dictionary<string, string> result = new Dictionary<string, string>();
                result.Add("Response", null);
                result.Add("Email", null);

                SmtpClient client = new SmtpClient();
                client.Port = HomeController.MAIL_PORT_NUMBER;
                client.DeliveryMethod = SmtpDeliveryMethod.Network;
                client.UseDefaultCredentials = false;
                client.Host = HomeController.MAIL_HOST_SERVER;


                MailMessage mail = new MailMessage();
                mail.From = new MailAddress(HomeController.MAIL_ADDRESS_SENDER);
                mail.IsBodyHtml = true;
                mail.Body = email.Body;
                mail.Body += "<br/> <span style='color:#e75114;'>" +
                    "Cet email a été renvoyé manuellement; la date " + DateTime.Now + " n'est donc pas la date réelle d'émission de ce rapport. " +
                    "Prière d'en tenir compte dans vos diverses interprétations des résultats qu'il fournit.</span><br />";
                mail.Attachments.Add(new Attachment(@email.Report.ResultPath));
                mail.Subject = "(Renvoi) " + email.Subject;
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
                if (ModelState.IsValid)
                {
                    db.Entry(email).State = System.Data.Entity.EntityState.Modified;
                    db.SaveChanges();
                }
                result["Response"] = "Votre email a été correctement envoyé aux personnes sélectionnées.";
                McoUtilities.General_Logging(new Exception("...."), "ReSend Email", 3);
                return Json(result, JsonRequestBehavior.AllowGet);
            }
            catch (Exception exception)
            {
                McoUtilities.General_Logging(exception, "ReSend Email", 0);
                Dictionary<string, string> result = new Dictionary<string, string>();
                result.Add("Response", "Une erreur est survenue durant l'envoi du message");
                result.Add("Error", exception.Message);
                return Json(result, JsonRequestBehavior.AllowGet);
            }
        }

        public JsonResult AutoSend(int id)
        {
            try
            {
                Email email = db.Emails.Find(id);
                if (email == null)
                {
                    return Json(HttpNotFound(), JsonRequestBehavior.AllowGet);
                }

                Dictionary<string, string> result = new Dictionary<string, string>();
                result.Add("Response", null);
                result.Add("Email", null);

                if (ModelState.IsValid)
                {
                    db.Entry(email).State = System.Data.Entity.EntityState.Modified;
                    db.SaveChanges();
                    Report report = db.Reports.Find(email.Report.Id);
                    TimeSpan Duration = DateTime.Now.Subtract(report.DateTime);
                    report.Duration = Duration;
                    if (ModelState.IsValid)
                    {
                        db.Entry(report).State = System.Data.Entity.EntityState.Modified;
                        db.SaveChanges();
                        result["Email"] = "Changement(s) sauvegardé(s)";
                    }
                }

                SmtpClient client = new SmtpClient();
                client.Port = HomeController.MAIL_PORT_NUMBER;
                client.DeliveryMethod = SmtpDeliveryMethod.Network;
                client.UseDefaultCredentials = false;
                client.Host = HomeController.MAIL_HOST_SERVER;


                MailMessage mail = new MailMessage();
                mail.From = new MailAddress(HomeController.MAIL_ADDRESS_SENDER);
                mail.IsBodyHtml = true;
                mail.Body = email.Body;
                mail.Attachments.Add(new Attachment(@email.Report.ResultPath));
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
                if (ModelState.IsValid)
                {
                    db.Entry(email).State = System.Data.Entity.EntityState.Modified;
                    db.SaveChanges();
                }
                result["Response"] = "Votre email a été correctement envoyé aux personnes sélectionnées.";
                McoUtilities.General_Logging(new Exception("...."), "Auto Email", 3);
                return Json(result, JsonRequestBehavior.AllowGet);
            }
            catch (Exception exception)
            {
                McoUtilities.General_Logging(exception, "Auto Email", 0);
                Dictionary<string, string> result = new Dictionary<string, string>();
                result.Add("Response", "Une erreur est survenue durant l'envoi du message");
                result.Add("Error", exception.Message);
                return Json(result, JsonRequestBehavior.AllowGet);
            }
        }

        public Email SetRecipients(Email email, string module)
        {
            if (email == null)
            {
                return null;
            }
            ICollection<Recipient> recipients = db.Recipients.Where(rec => rec.Module == module)
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

        protected override void Dispose(bool disposing)
        {
            db.Dispose();
            base.Dispose(disposing);
        }
    }
}