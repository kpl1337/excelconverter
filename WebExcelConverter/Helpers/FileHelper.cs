using System;
using System.IO;
using ExcelToJsonWeb.Helpers;

namespace WebExcelConverter.Helpers
{
    public static class FileHelper
    {
        /// <summary>
        /// Method saves uploaded excel file
        /// </summary>
        /// <param name="fileStream">Stream with excel</param>
        /// <returns>Path to saved file</returns>
        public static string SaveUploadedFile(Stream fileStream)
        {
            string tempFilePath = Path.Combine(Path.GetTempPath(), Path.GetRandomFileName() + Constants.ExcelFileExtension);

            using (var stream = new FileStream(tempFilePath, FileMode.Create))
            {
                fileStream.CopyTo(stream);
            }

            return tempFilePath;
        }

        /// <summary>
        /// Compose the outputh json path
        /// </summary>
        /// <param name="originalFileName">Roginal file name</param>
        /// <returns></returns>
        public static string GetJsonOutputPath(string originalFileName)
        {
            string jsonFileName = Path.GetFileNameWithoutExtension(originalFileName) + Constants.JsonFileExtension;
            return Path.Combine(Path.GetTempPath(), jsonFileName);
        }
    }
}

