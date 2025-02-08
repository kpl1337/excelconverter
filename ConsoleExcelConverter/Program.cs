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

        IExcelConverter converter = new JsonExcelConverter();

        foreach (string filePath in excelFiles)
        {
            ProcessFile(filePath, converter);
        }

        Console.WriteLine("All conversions completed.");
    }

    static void ProcessFile(string filePath, IExcelConverter converter)
    {
        string fileName = Path.GetFileName(filePath);
        string outputFileName = Path.GetFileNameWithoutExtension(filePath) + ".json";

        try
        {
            Console.WriteLine($"Processing: {fileName}");
            string result = converter.Convert(filePath);
            File.WriteAllText(outputFileName, result);
            Console.WriteLine($"Successfully converted '{fileName}' to '{outputFileName}'");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error processing '{fileName}': {ex.Message}");
        }
    }
}


