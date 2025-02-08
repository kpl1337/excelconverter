using Microsoft.AspNetCore.Mvc;
using ExcelToJsonWeb.Helpers;
using WebExcelConverter.Models;
using WebExcelConverter.Services;

namespace WebExcelConverter.Controllers
{
    public class HomeController : Controller
    {
        private readonly ExcelService _excelService;

        public HomeController()
        {
            _excelService = new ExcelService();
        }

        public IActionResult Index()
        {
            return View();
        }

        [HttpPost]
        public IActionResult UploadFile(UploadModel model)
        {
            if (model.File == null || model.File.Length == 0)
            {
                ViewBag.Message = Constants.FileNotSelected;
                return View("Index");
            }

            if (Path.GetExtension(model.File.FileName).ToLower() != Constants.ExcelFileExtension)
            {
                ViewBag.Message = Constants.InvalidFileType;
                return View("Index");
            }

            try
            {
                string outputFilePath = _excelService.ConvertExcelToJson(model.File.OpenReadStream(), model.File.FileName);

                ViewBag.JsonFilePath = Path.GetFileName(outputFilePath);
                ViewBag.Message = Constants.ConversionSuccess;

                return View("Index");
            }
            catch (System.Exception ex)
            {
                ViewBag.Message = $"{Constants.ErrorPrefix} {ex.Message}";
                return View("Index");
            }
        }

        public IActionResult DownloadFile(string fileName)
        {
            string filePath = Path.Combine(Path.GetTempPath(), fileName);
            if (System.IO.File.Exists(filePath))
            {
                byte[] fileBytes = System.IO.File.ReadAllBytes(filePath);
                return File(fileBytes, "application/json", fileName);
            }
            return NotFound();
        }
    }
}

