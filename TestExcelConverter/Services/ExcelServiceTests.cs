using FluentAssertions;
using ClosedXML.Excel;
using WebExcelConverter.Services;

namespace TestExcelConverter.Services
{
    public class ExcelServiceTests
    {
        private readonly ExcelService _excelService;

        public ExcelServiceTests()
        {
            _excelService = new ExcelService();
        }

        private string CreateValidTestExcelFile()
        {
            string tempFilePath = Path.Combine(Path.GetTempPath(), Path.GetRandomFileName() + ".xlsx");

            using (var workbook = new XLWorkbook())
            {
                var worksheet = workbook.Worksheets.Add("Test");
                worksheet.Cell(1, 1).Value = "Název klienta";
                worksheet.Cell(1, 2).Value = "IČ klienta";
                worksheet.Cell(1, 3).Value = "Zakázka";
                worksheet.Cell(2, 1).Value = "Test Klient";
                worksheet.Cell(2, 2).Value = "12345678";
                worksheet.Cell(2, 3).Value = "ZK001";

                workbook.SaveAs(tempFilePath);
            }

            return tempFilePath;
        }

        [Fact]
        public void ConvertExcelToJson_ShouldReturnJsonFilePath()
        {
            // Arrange
            string tempExcelFile = CreateValidTestExcelFile();
            string originalFileName = "TestFile.xlsx";
            string resultFilePath;

            using (var stream = new FileStream(tempExcelFile, FileMode.Open, FileAccess.Read, FileShare.Read))
            {
                // Act
                resultFilePath = _excelService.ConvertExcelToJson(stream, originalFileName);
            }

            // Assert
            resultFilePath.Should().NotBeNullOrWhiteSpace();
            resultFilePath.Should().EndWith(".json");
            File.Exists(resultFilePath).Should().BeTrue();

            // Cleanup
            try
            {
                if (File.Exists(resultFilePath)) File.Delete(resultFilePath);
                if (File.Exists(tempExcelFile)) File.Delete(tempExcelFile);
            }
            catch (IOException ex)
            {
                Console.WriteLine($"Warning: Soubor nelze smazat - {ex.Message}");
            }
        }
    }
}


