using OfficeOpenXml;
using System.IO;
using System.Text.Json;


const string CONFIG_PATH = "config.json";

if (!File.Exists("config.json"))
{
    Console.WriteLine("No config file");
    return;
}

Config config = JsonSerializer.Deserialize<Config>(File.ReadAllText(CONFIG_PATH));

var inputFilePath = config.InputPath;
var outputFilePath = config.OutputPath;

ExcelPackage.LicenseContext = LicenseContext.NonCommercial; // or LicenseContext.Commercial if you have a commercial license


using (var inputPackage = new ExcelPackage(new FileInfo(inputFilePath)))
{
    ExcelWorksheet inputWorksheet = inputPackage.Workbook.Worksheets[0];

    using (var outputPackage = new ExcelPackage())
    {
        ExcelWorksheet outputWorksheet = outputPackage.Workbook.Worksheets.Add("FilteredData");
        
        Console.WriteLine($"Found {inputWorksheet.Dimension.End.Row} rows");

        int outputRowIndex = 1;

        for (int i = 1; i < config.StartingFrom; i++)
        {
            outputWorksheet.Cells[outputRowIndex, 1, outputRowIndex, inputWorksheet.Dimension.End.Column].Value =
                inputWorksheet.Cells[i, 1, i, inputWorksheet.Dimension.End.Column].Value;

            outputRowIndex++;
        }

        for (int i = config.StartingFrom; i <= inputWorksheet.Dimension.End.Row; i++)
        {
            var cellValue = inputWorksheet.Cells[i, config.ColumnId].GetValue<string>();

            if (Array.Exists(config.FilterCriteria, criteria => criteria.Equals(cellValue, StringComparison.OrdinalIgnoreCase)))
            {
                outputWorksheet.Cells[outputRowIndex, 1, outputRowIndex, inputWorksheet.Dimension.End.Column].Value =
                    inputWorksheet.Cells[i, 1, i, inputWorksheet.Dimension.End.Column].Value;

                outputRowIndex++;
            }
        }

        outputPackage.SaveAs(new FileInfo(outputFilePath));

        Console.WriteLine($"Filtered data has been saved to {outputFilePath}");
    }
}

// Class to hold configuration data from config.json
class Config
{
    public string[] FilterCriteria { get; set; }
    public int ColumnId { get; set; }
    public int StartingFrom { get; set; }
    public string InputPath { get; set; }
    public string OutputPath { get; set; }
}

