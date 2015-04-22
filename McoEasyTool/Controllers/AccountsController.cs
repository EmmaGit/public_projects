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
    public class AccountsController : Controller
    {
        private DataModelContainer db = new DataModelContainer();

        public ActionResult DisplayAccounts()
        {
            Initialize();
            return View(db.Accounts.OrderBy(acc => acc.DisplayName));
        }

        public bool Initialize()
        {
            IQueryable<Account> default_general_username = db.Accounts
                .Where(acc => acc.DisplayName == HomeController.DEFAULT_DOMAIN_AND_USERNAME_IMPERSONNATION);
            if (default_general_username.Count() == 0)
            {
                Account account = db.Accounts.Create();
                account.DisplayName = HomeController.DEFAULT_DOMAIN_AND_USERNAME_IMPERSONNATION.ToUpper();
                account.Domain = HomeController.DEFAULT_DOMAIN_IMPERSONNATION.ToUpper();
                account.Username = HomeController.DEFAULT_USERNAME_IMPERSONNATION.ToUpper();
                account.Password = HomeController.DEFAULT_PASSWORD_IMPERSONNATION;
                account.IsSystem = true;
                if (ModelState.IsValid)
                {
                    db.Accounts.Add(account);
                    db.SaveChanges();
                    McoUtilities.General_Logging(new Exception("...."), "Initialize " + account.DisplayName, 2);
                }
            }

            IQueryable<Account> space_general_username = db.Accounts
                .Where(acc => acc.DisplayName == HomeController.SPACE_DOMAIN_AND_USERNAME_IMPERSONNATION);
            if (space_general_username.Count() == 0)
            {
                Account account = db.Accounts.Create();
                account.DisplayName = HomeController.SPACE_DOMAIN_AND_USERNAME_IMPERSONNATION.ToUpper();
                account.Domain = HomeController.SPACE_DOMAIN_IMPERSONNATION.ToUpper();
                account.Username = HomeController.SPACE_USERNAME_IMPERSONNATION.ToUpper();
                account.Password = HomeController.SPACE_PASSWORD_IMPERSONNATION;
                account.IsSystem = true;
                if (ModelState.IsValid)
                {
                    db.Accounts.Add(account);
                    db.SaveChanges();
                    McoUtilities.General_Logging(new Exception("...."), "Initialize " + account.DisplayName, 2);
                }
            }

            return true;
        }

        [HttpPost]
        public string AddAccount()
        {
            try
            {
                string domain = "";
                string username = Request.Form["username"];
                string password = Request.Form["password"];
                if (username == null || username.Trim() == "" || username.IndexOf(@"\") == -1)
                {
                    return @"Veuillez renseigner un nom d'utilisateur sous la forme Domain\username";
                }
                if (password == null || password.Trim() == "")
                {
                    return "Les mots de passe vides ne sont pas autorisés.";
                }
                string[] infos = username.Split(new string[] { @"\" }, StringSplitOptions.RemoveEmptyEntries);
                if (infos.Length != 2)
                {
                    return @"Veuillez renseigner un nom d'utilisateur sous la forme Domain\username";
                }
                domain = infos[0];
                username = infos[1];
                Account account = db.Accounts.Create();
                account.Domain = domain.Trim().ToUpper();
                account.Username = username.Trim().ToUpper();
                account.Password = McoUtilities.Encrypt(password);
                account.DisplayName = account.Domain + "\\" + account.Username;

                account.IsSystem = false;
                if (!McoUtilities.IsValidLoginPassword(account.DisplayName, account.Password))
                {
                    //    return "Erreur, le couple user/password ne correpond pas.";
                }
                if (Exists(account))
                {
                    return "Ce compte d'utilisateur existe déjà en base.";
                }
                if (ModelState.IsValid)
                {
                    db.Accounts.Add(account);
                    db.SaveChanges();
                    McoUtilities.General_Logging(new Exception("...."), "AddAccount " + account.DisplayName, 2, User.Identity.Name);
                    return "Ce compte a été rajouté à la liste.";
                }
                McoUtilities.General_Logging(new Exception("...."), "AddAccount " + account.DisplayName, 2, User.Identity.Name);
                return "Erreur lors de l'ajout du compte";
            }
            catch (Exception exception)
            {
                McoUtilities.General_Logging(exception, "AddAccount");
                return "Erreur lors de l'ajout du compte";
            }
        }

        [HttpPost]
        public string EditAccount(int id)
        {
            try
            {
                Account account = db.Accounts.Find(id);
                if (account == null)
                {
                    return "Ce compte d'utilisateur n'existe pas en mémoire";
                }
                if (account.IsSystem)
                {
                    return "Il s'agit d'un compte système; pour le modifier veuillez-contacter un administrateur du site";
                }
                string password = Request.Form["password"];
                if (password == null || password.Trim() == "")
                {
                    return "Les mots de passe vides ne sont pas autorisés.";
                }
                account.Password = McoUtilities.Encrypt(password);
                if (!McoUtilities.IsValidLoginPassword(account.DisplayName, account.Password))
                {
                    //    return "Erreur, le couple user/password ne correpond pas.";
                }
                if (ModelState.IsValid)
                {
                    db.Entry(account).State = System.Data.Entity.EntityState.Modified;
                    db.SaveChanges();
                    McoUtilities.General_Logging(new Exception("...."), "EditAccount " + account.DisplayName, 2, User.Identity.Name);
                    account.DisplayName = account.Domain + "\\" + account.Username;
                    return "Le mot de passe du compte a été mis à jour";
                }
                McoUtilities.General_Logging(new Exception("...."), "EditAccount " + account.DisplayName, 2, User.Identity.Name);
                return "Erreur lors de la mise à jour du compte";
            }
            catch (Exception exception)
            {
                McoUtilities.General_Logging(exception, "EditAccount");
                return "Erreur lors de la mise à jour du compte";
            }
        }

        public string DeleteAccount(int id)
        {
            Account account = db.Accounts.Find(id);
            string username = account.DisplayName;
            if (account == null)
            {
                return "Compte non trouvé en bases";
            }
            if (account.IsSystem)
            {
                return "La suppression des comptes systèmes est impossible, contactez un administrateur du site";
            }
            UpdatePoolsAccounts(account);
            UpdateSpaceServersAccounts(account);
            db.Accounts.Remove(account);
            db.SaveChanges();
            McoUtilities.General_Logging(new Exception("...."), "DeleteAccount " + username, 2, User.Identity.Name);
            return "Le compte a été supprimé";
        }

        public bool UpdatePoolsAccounts(Account account)
        {
            List<Pool> pools = db.Pools.Where(poo => poo.CheckAccount == account.DisplayName).ToList();
            foreach (Pool pool in pools)
            {
                pool.CheckAccount = HomeController.DEFAULT_DOMAIN_AND_USERNAME_IMPERSONNATION;
                if (ModelState.IsValid)
                {
                    db.Entry(pool).State = System.Data.Entity.EntityState.Modified;
                    db.SaveChanges();
                    McoUtilities.Specific_Logging(new Exception("...."),
                        "UpdatePoolsAccount " + pool.Name + " CheckAccount", HomeController.BESR_MODULE, 2);
                }
            }
            pools = new List<Pool>();
            pools = db.Pools.Where(poo => poo.ExecutionAccount == account.DisplayName).ToList();
            foreach (Pool pool in pools)
            {
                pool.ExecutionAccount = HomeController.DEFAULT_DOMAIN_AND_USERNAME_IMPERSONNATION;
                if (ModelState.IsValid)
                {
                    db.Entry(pool).State = System.Data.Entity.EntityState.Modified;
                    db.SaveChanges();
                    McoUtilities.Specific_Logging(new Exception("...."),
                        "UpdatePoolsAccount " + pool.Name + " ExecutionAccount", HomeController.BESR_MODULE, 2);
                }
            }
            return true;
        }

        public bool UpdateSpaceServersAccounts(Account account)
        {
            List<SpaceServer> servers = db.SpaceServers.Where(poo => poo.CheckAccount == account.DisplayName).ToList();
            foreach (SpaceServer server in servers)
            {
                server.CheckAccount = HomeController.SPACE_DOMAIN_AND_USERNAME_IMPERSONNATION;
                if (ModelState.IsValid)
                {
                    db.Entry(server).State = System.Data.Entity.EntityState.Modified;
                    db.SaveChanges();
                    McoUtilities.Specific_Logging(new Exception("...."),
                        "UpdateSpaceServersAccounts " + server.Name + " CheckAccount", HomeController.SPACE_MODULE, 2);
                }
            }
            servers = new List<SpaceServer>();
            servers = db.SpaceServers.Where(poo => poo.ExecutionAccount == account.DisplayName).ToList();
            foreach (SpaceServer server in servers)
            {
                server.ExecutionAccount = HomeController.SPACE_DOMAIN_AND_USERNAME_IMPERSONNATION;
                if (ModelState.IsValid)
                {
                    db.Entry(server).State = System.Data.Entity.EntityState.Modified;
                    db.SaveChanges();
                    McoUtilities.Specific_Logging(new Exception("...."),
                        "UpdateSpaceServersAccounts " + server.Name + " ExecutionAccount", HomeController.SPACE_MODULE, 2);
                }
            }
            return true;
        }

        public bool IsAlready(string displayname)
        {
            IQueryable<Account> accounts = db.Accounts.Where(acc => acc.DisplayName == displayname);
            if (accounts.Count() == 0)
            {
                return false;
            }
            return true;
        }

        public bool Exists(Account account)
        {
            if (!IsAlready(account.DisplayName))
            {
                IQueryable<Account> accounts = db.Accounts.Where(acc => acc.Domain == account.Domain)
                    .Where(acc => acc.Username == account.Username);
                if (accounts.Count() == 0)
                {
                    return false;
                }
                return true;
            }
            return true;
        }

        public List<Account> GetAccountsList()
        {
            Initialize();
            List<Account> accounts = db.Accounts.ToList();
            return accounts;
        }

        public string GetStringedAccountsList()
        {
            string options = "";
            Initialize();
            List<Account> accounts = db.Accounts.ToList();
            foreach (Account account in accounts)
            {
                options += "<option val='" + account.DisplayName +
                    "'>" + account.DisplayName + "</option>";
            }
            return options;
        }

        public Account GetAccountByDisplayName(string username)
        {
            Account account = db.Accounts.Where(acc => acc.DisplayName.ToUpper() == username.ToUpper())
                .FirstOrDefault();
            return account;
        }
    }
}
