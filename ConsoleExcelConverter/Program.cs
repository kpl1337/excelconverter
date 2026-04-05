using ExcelConverterLibrary;

class Program
{
    static void Main()
    {
        string currentDirectory = Directory.GetCurrentDirectory();
        string[] excelFiles = Directory.GetFiles(currentDirectory, "*.xlsx");

        if (excelFiles.Length == 0)
        {
            Console.WriteLine("No Excel (.xlsx) files found in the current directory.");
            return;
        }

        Console.WriteLine($"Found {excelFiles.Length} Excel file(s). Converting...");

        IExcelConverter jsonConverter = new JsonExcelConverter();
        IExcelConverter xmlConverter = new XmlExcelConverter();

        foreach (string filePath in excelFiles)
        {
            ProcessFile(filePath, jsonConverter);
            ProcessFile(filePath, xmlConverter);
        }

        Console.WriteLine("All conversions completed.");
    }

    static void ProcessFile(string filePath, IExcelConverter converter)
    {
        string fileName = Path.GetFileName(filePath);
        string outputFileNameJson = Path.GetFileNameWithoutExtension(filePath) + ".json";
        string outputFileNameXml = Path.GetFileNameWithoutExtension(filePath) + ".xml";

        try
        {
            Console.WriteLine($"Processing: {fileName}");
            string jsonResult = converter.Convert(filePath);
            string xmlResult = converter.Convert(filePath);
            File.WriteAllText(outputFileNameJson, jsonResult);
            File.WriteAllText(outputFileNameXml, xmlResult);
            Console.WriteLine($"Successfully converted '{fileName}' to '{outputFileNameJson}'");
            Console.WriteLine($"Successfully converted '{fileName}' to '{outputFileNameXml}'");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error processing '{fileName}': {ex.Message}");
        }
    }
}


