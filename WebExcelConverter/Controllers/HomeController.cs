using Microsoft.AspNetCore.Mvc;
using ExcelToJsonWeb.Helpers;
using WebExcelConverter.Models;
using WebExcelConverter.Services;
using DocumentFormat.OpenXml.EMMA;

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
                if (model.Format == "json")
                {
                    string outputFilePath = _excelService.ConvertExcelToJson(model.File.OpenReadStream(), model.File.FileName);

                    ViewBag.JsonFilePath = Path.GetFileName(outputFilePath);
                    ViewBag.Message = Constants.ConversionSuccess;

                    return View("Index");
                } 
                else //if (model.Format == "xml") // all if-statement paths have to return a value, so, this wouldn't work 
                {
                    string outputFilePath = _excelService.ConvertExcelToXml(model.File.OpenReadStream(), model.File.FileName);

                    ViewBag.XmlFilePath = Path.GetFileName(outputFilePath);
                    ViewBag.Message = Constants.ConversionSuccess;

                    return View("Index");
                }
            }
            
            catch (System.Exception ex)
            {
                ViewBag.Message = $"{Constants.ErrorPrefix} {ex.Message}";
                return View("Index");
            }
        }

        /// <summary>
        /// Method downloads converted file
        /// </summary>
        /// <param name="fileName">Output file name</param>
        /// <param name="format">Format of the output file</param>
        /// <returns>returns file</returns>

        public IActionResult DownloadFile(string fileName, string format)
        {
            string filePath = Path.Combine(Path.GetTempPath(), fileName);

            if (System.IO.File.Exists(filePath))
            {
                byte[] fileBytes = System.IO.File.ReadAllBytes(filePath);
                
                // already defined in the function argument, just so we don't have to
                // make a whole another function, or logic
                return File(fileBytes, "application/" + format, fileName);
            }

            return NotFound();
        }
    }
}
