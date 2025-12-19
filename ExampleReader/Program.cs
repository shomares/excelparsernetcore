using ExcelReader.src.Implementation;

using var simple = new ReaderExcelSimple();

var result = await simple.GetNextRowAsync("C:\\Users\\MX05711629\\OneDrive - Coca-Cola FEMSA\\Documentos\\ExampleBig.xlsx", "sheet1");

Console.WriteLine("Done");
