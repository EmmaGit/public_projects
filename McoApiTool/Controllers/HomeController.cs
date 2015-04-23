using System.Web.Mvc;
using MODELS = McoApiTool.Models;

namespace McoApiTool.Controllers
{
    public class HomeController : Controller
    {
        public ActionResult Index()
        {
            MODELS.User user = UtilitiesController.HostIfExists(User.Identity.Name);
            if (user != null)
            {
                return RedirectToAction("Home", "Profile", new { id = user.Id });
            }
            ViewBag.Title = "Sign Up now please";
         
            //LoginPage
            return View();
        }

    }
}
