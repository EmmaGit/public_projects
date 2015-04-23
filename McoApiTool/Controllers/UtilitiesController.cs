using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using System.DirectoryServices;
using McoApiTool.Models;

namespace McoApiTool.Controllers
{
    public class UtilitiesController : Controller
    {
        private static McoApiToolContext db = new McoApiToolContext();
        public static User HostIfExists(string hostname)
        {
            if (db.Users.Where(u => u.Hostname == hostname.ToUpper()).Count() > 0)
            {
                return db.Users.Where(u => u.Hostname == hostname.ToUpper()).FirstOrDefault();
            }
            return null;
        }

        public List<string> GetHostnames() 
        {
            DirectoryEntry root = new DirectoryEntry("WinNT:");
            List<string> hostnames = new List<string>();
            foreach (DirectoryEntry computers in root.Children)
            {
                foreach (DirectoryEntry computer in computers.Children)
                {
                    if (computer.Name != "Schema")
                    {
                        hostnames.Add(computer.Name);
                    }
                }
            }
            return hostnames;
        }
	}
}