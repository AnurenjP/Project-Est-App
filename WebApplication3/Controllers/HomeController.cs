using Microsoft.AspNetCore.Mvc;
using System.Diagnostics;
using ProjectEstimationApp.Models;

namespace ProjectEstimationApp.Controllers
{
    public class HomeController : Controller
    {
        private readonly ILogger<HomeController> _logger;

        public HomeController(ILogger<HomeController> logger)
        {
            _logger = logger;
        }

        public IActionResult Index()
        {
            return View();
        }

        [HttpPost]
        public IActionResult SubmitForm(ProjectEstimation model)
        {
            // Store the form data in TempData to pass it to the next page
            TempData["ProjectEstimation"] = model;

            // Redirect to the next page
            return RedirectToAction("Estpage");
        }

        public IActionResult Estpage()
        {
            var model = TempData["ProjectEstimation"] as ProjectEstimation;
            return View(model);
        }

        public IActionResult Resource()
        {
            return View();
        }

        public IActionResult Sampletimeline()
        {
            return View();
        }

        public IActionResult smbud()
        {
            return View();
        }

        public IActionResult Privacy()
        {
            return View();
        }

        public IActionResult Homepage()
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