using Microsoft.AspNetCore.Mvc;
using System.Diagnostics;
using ProjectEstimationApp.Models;

namespace ProjectEstimationApp.Controllers
{
    public class HomeController : Controller
    {
        [HttpPost]
        public IActionResult SubmitForm(string projectType, string complexity, int hours)
        {
            // Store the form data in TempData to pass it to the next page
            TempData["ProjectType"] = projectType;
            TempData["Complexity"] = complexity;
            TempData["Hours"] = hours;

            // Redirect to the next page
            return RedirectToAction("Index");
        }
        private readonly ILogger<HomeController> _logger;

        public HomeController(ILogger<HomeController> logger)
        {
            _logger = logger;
        }

        public ActionResult Index()
        {
            return View();
        }
        public ActionResult Index2()
        {
            return View();
        }
        public ActionResult Index3()
        {
            return View();
        }
        public ActionResult smbud()
        {
            return View();
        }



        public ActionResult Privacy()
        {
            return View();
        }

        public ActionResult Homepage()
        {
            return View();
        }


        [ResponseCache(Duration = 0, Location = ResponseCacheLocation.None, NoStore = true)]
        public IActionResult Error()
        {
            return View(new ErrorViewModel { RequestId = Activity.Current?.Id ?? HttpContext.TraceIdentifier });
        }
    }
}
