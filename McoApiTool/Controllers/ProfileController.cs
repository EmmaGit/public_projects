using System.Threading.Tasks;
using System.Web.Http;
using System.Web.Mvc;
using MODELS = McoApiTool.Models;

namespace McoApiTool.Controllers
{
    public class ProfileController : Controller
    {
        private static MODELS.McoApiToolContext db = new MODELS.McoApiToolContext();
        private UsersController users_controller = new UsersController();
        // GET: Profile
        public ActionResult Home(int id)
        {
            MODELS.User user = db.Users.Find(id);
            if (user == null)
            {
                return HttpNotFound();
            }
            //response.
            ViewData["Id"] = id;
            //ViewData["Username"] = response.Result;
            ViewBag.Title = "Home Page";
            return View(user);
        }
    }
}