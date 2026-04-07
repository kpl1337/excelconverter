using System;
using System.IO;
using ExcelConverterLibrary;
using ExcelToJsonWeb.Helpers;
using WebExcelConverter.Helpers;

namespace WebExcelConverter.Services
{
    public class ExcelService
    {
        // replaced in each method
        //public ExcelService()
        //{
        //    _converter = new JsonExcelConverter();
        //}

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
                IExcelConverter jsonConverter = new JsonExcelConverter();
                string jsonResult = jsonConverter.Convert(tempFilePath);
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

        /// <summary>
        /// Converts Excel to Xml using the library
        /// </summary>
        /// <param name="fileStream">Stream with excel file</param>
        /// <param name="originalFileName">Original file name</param>
        /// <returns></returns>
        public string ConvertExcelToXml(Stream fileStream, string originalFileName)
        {
            string tempFilePath = FileHelper.SaveUploadedFile(fileStream);

            try
            {
                IExcelConverter xmlConverter = new XmlExcelConverter();
                string xmlResult = xmlConverter.Convert(tempFilePath);
                string outputPath = FileHelper.GetXmlOutputPath(originalFileName);
                
                File.WriteAllText(outputPath, xmlResult);
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

