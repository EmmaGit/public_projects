using System.Data;
using System.Linq;
using System.Web.Mvc;
using System;

namespace McoEasyTool.Controllers
{
    [AllowAnonymous]
    public class RecipientsController : Controller
    {
        private DataModelContainer db = new DataModelContainer();

        [HttpPost]
        public string Create()
        {
            try
            {
                string relative = Request.Form["relative"];
                string absolute = Request.Form["absolute"];
                string name = Request.Form["name"];
                string module = Request.Form["module"];
                string included = Request.Form["included"];

                Recipient recipient = db.Recipients.Create();
                recipient.RelativeAddress = relative;
                recipient.AbsoluteAddress = absolute;
                recipient.Name = name;
                recipient.Module = module;
                recipient.Included = (included == null) ? true : (included == "true");
                if (ModelState.IsValid)
                {
                    db.Recipients.Add(recipient);
                    db.SaveChanges();
                    McoUtilities.General_Logging(new Exception("...."), "Create " + module + " " + recipient.AbsoluteAddress, 3, User.Identity.Name);
                    return "Le destinataire a été correctement rajouté à la liste.";
                }
                McoUtilities.General_Logging(new Exception("...."), "Create " + module + " " + recipient.AbsoluteAddress, 2, User.Identity.Name);
                return "Erreur lors de l'ajout du destinataire";
            }
            catch (Exception exception)
            {
                McoUtilities.General_Logging(exception, "Create Recipient", 0);
                return "Erreur lors de l'ajout du destinataire";
            }
        }

        [HttpPost]
        public string Edit(int id)
        {
            try
            {
                string relative = Request.Form["relative"];
                string absolute = Request.Form["absolute"];
                string name = Request.Form["name"];
                string module = Request.Form["module"];
                string included = Request.Form["included"];

                Recipient recipient = db.Recipients.Find(id);
                recipient.RelativeAddress = relative;
                recipient.AbsoluteAddress = absolute;
                recipient.Name = name;
                recipient.Module = module;
                if (included != null)
                {
                    recipient.Included = (included == "true");
                }

                if (ModelState.IsValid)
                {
                    db.Entry(recipient).State = System.Data.Entity.EntityState.Modified;
                    db.SaveChanges();
                    McoUtilities.General_Logging(new Exception("...."), "Edit " + module + " " + recipient.AbsoluteAddress, 3, User.Identity.Name);
                    return "Les modifications ont été enrégistrées.";
                }
                McoUtilities.General_Logging(new Exception("...."), "Edit " + module + " " + recipient.AbsoluteAddress, 2, User.Identity.Name);
                return "Erreur lors de la modification du destinataire";
            }
            catch (Exception exception)
            {
                McoUtilities.General_Logging(exception, "Edit Recipient", 0);
                return "Erreur lors de la modification du destinataire";
            }
        }

        public string Delete(int id)
        {
            Recipient recipient = db.Recipients.Find(id);
            string module = recipient.Module;
            db.Recipients.Remove(recipient);
            db.SaveChanges();
            McoUtilities.General_Logging(new Exception("...."), "Delete " + module + " " + recipient.AbsoluteAddress, 3, User.Identity.Name);
            return "Ce destinataire a été supprimé de la liste";
        }

        protected override void Dispose(bool disposing)
        {
            db.Dispose();
            base.Dispose(disposing);
        }
    }
}