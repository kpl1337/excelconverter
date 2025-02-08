using System;
using System.IO;
using ExcelConverterLibrary;
using ExcelToJsonWeb.Helpers;
using WebExcelConverter.Helpers;

namespace WebExcelConverter.Services
{
    public class ExcelService
    {
        private readonly IExcelConverter _converter;

        public ExcelService()
        {
            _converter = new JsonExcelConverter();
        }

        /// <summary>
        /// Converts Excel to Json using the library
        /// </summary>
        /// <param name="fileStream">Stream with excel file</param>
        /// <param name="originalFileName">Original file name</param>
        /// <returns></returns>
        public string ConvertExcelToJson(Stream fileStream, string originalFileName)
        {
            string tempFilePath = FileHelper.SaveUploadedFile(fileStream);

            try
            {
                string jsonResult = _converter.Convert(tempFilePath);
                string outputPath = FileHelper.GetJsonOutputPath(originalFileName);
                File.WriteAllText(outputPath, jsonResult);
                return outputPath;
            }
            finally
            {
                try
                {
                    if (File.Exists(tempFilePath))
                    {
                        File.Delete(tempFilePath);
                    }
                }
                catch (IOException ex)
                {
                    Console.WriteLine($"Warning: Could not delete temp file {tempFilePath}. Error: {ex.Message}");
                }
            }
        }
    }
}

