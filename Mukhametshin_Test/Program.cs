using OfficeOpenXml;

const string fileName = "ТЗ.Transmittal.xlsx";
const int listNumber = 0;
const int dataStartRow = 2;
const int nameColumnNumber = 1;
const int pathColumnNumber = 2;
const int extensionColumnNumber = 3;

ExcelPackage.License.SetNonCommercialPersonal("Danil");
var data = ParseFile(fileName);
BuildFileStructure(data);
return;

static void BuildFileStructure(List<Data> dataList)
{
    foreach (var data in dataList)
    {
        Directory.CreateDirectory(data.Path);

        var filePath = data.FullPath;
        if (File.Exists(filePath) == false)
        {
            File.Create(filePath).Dispose();
            Console.WriteLine("Создан файл {0}", filePath);
        }
        else
        {
            Console.WriteLine("Файл уже существует {0}", filePath);
        }    
    }
}

static List<Data> ParseFile(string excelPath)
{
    using var excel = new ExcelPackage(new FileInfo(excelPath));
    if (excel.Workbook.Worksheets.Count <= listNumber)
    {
        return [];
    }
    
    var sheet = excel.Workbook.Worksheets[listNumber];
    var result = new List<Data>();

    for (var row = dataStartRow; row <= sheet.Dimension.Rows; row++)
    {
        var name = sheet.Cells[row, nameColumnNumber].Value?.ToString()?.Trim();
        var path = sheet.Cells[row, pathColumnNumber].Value?.ToString()?.Trim();
        var extension = sheet.Cells[row, extensionColumnNumber].Value?.ToString()?.Trim();

        if (string.IsNullOrWhiteSpace(name) ||
            string.IsNullOrWhiteSpace(path) ||
            string.IsNullOrWhiteSpace(extension))
        {
            Console.WriteLine("WRN: В строке {0} недостаточно данных!", row);
            continue;
        }

        result.Add(new Data
        {
            Name = name,
            Path = path,
            Extension = extension
        });
    }

    return result;
}

public class Data
{
    public required string Name { get; init; }
    public required string Path { get; init; }
    public required string Extension { get; init; }
    
    public string FullPath => System.IO.Path.Combine(Path, $"{Name}.{Extension}");
}