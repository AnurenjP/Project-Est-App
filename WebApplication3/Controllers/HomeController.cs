using Microsoft.AspNetCore.Mvc;
using System.Diagnostics;
using ProjectEstimationApp.Models;
using OfficeOpenXml;
using System.IO;

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

        public IActionResult Smbud()
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

        [HttpPost]
        public IActionResult GenerateFiles([FromBody] ProjectData projectData)
        {
            var excelFilePath = GenerateExcel(projectData);
            return Json(new { success = true, excelFilePath });
        }

        private string GenerateExcel(ProjectData projectData)
        {
            // Set the license context for EPPlus
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            var downloadsPath = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);
            var filePath = Path.Combine(downloadsPath, "Downloads", "ProjectEstimation.xlsx");

            try
            {
                using (var package = new ExcelPackage())
                {
                    var worksheet = package.Workbook.Worksheets.Add("Project Estimation");
                    worksheet.Cells[1, 1].Value = "Resource Type";
                    worksheet.Cells[1, 2].Value = "Cost";
                    worksheet.Cells[1, 3].Value = "Number of Resources";
                    worksheet.Cells[1, 4].Value = "Total";

                    int row = 2;
                    foreach (var resource in projectData.Resources)
                    {
                        worksheet.Cells[row, 1].Value = resource.Name;
                        worksheet.Cells[row, 2].Value = resource.Cost;
                        worksheet.Cells[row, 3].Value = resource.NumberOfResources;
                        worksheet.Cells[row, 4].Value = resource.Total;
                        row++;
                    }

                    worksheet.Cells[row, 1].Value = "Project Start Date";
                    worksheet.Cells[row, 2].Value = projectData.ProjectStartDate;

                    worksheet.Cells[row + 1, 1].Value = "Project End Date";
                    worksheet.Cells[row + 1, 2].Value = projectData.ProjectEndDate;

                    var additionalCostsSheet = package.Workbook.Worksheets.Add("Additional Costs");
                    additionalCostsSheet.Cells[1, 1].Value = "Name";
                    additionalCostsSheet.Cells[1, 2].Value = "Cost";
                    additionalCostsSheet.Cells[1, 3].Value = "Number of Resources";
                    additionalCostsSheet.Cells[1, 4].Value = "Total";

                    int additionalRow = 2;
                    foreach (var cost in projectData.AdditionalCosts)
                    {
                        additionalCostsSheet.Cells[additionalRow, 1].Value = cost.Name;
                        additionalCostsSheet.Cells[additionalRow, 2].Value = cost.Cost;
                        additionalCostsSheet.Cells[additionalRow, 3].Value = cost.NumberOfResources;
                        additionalCostsSheet.Cells[additionalRow, 4].Value = cost.Total;
                        additionalRow++;
                    }

                    package.SaveAs(new FileInfo(filePath));
                }
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error saving Excel file");
                throw;
            }

            return filePath;
        }

        [ResponseCache(Duration = 0, Location = ResponseCacheLocation.None, NoStore = true)]
        public IActionResult Error()
        {
            return View(new ErrorViewModel { RequestId = Activity.Current?.Id ?? HttpContext.TraceIdentifier });
        }
    }
}