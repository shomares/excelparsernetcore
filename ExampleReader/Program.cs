using ExcelReader.src.Implementation;
using System.Diagnostics;


using var simple = new ReaderExcelSimple();

Stopwatch stopwatch = Stopwatch.StartNew();
await simple.ReadFileAsync("C:\\Users\\MX05711629\\OneDrive - Coca-Cola FEMSA\\Documentos\\ExampleBig.xlsx", "sheet1");
var result = simple.GetNextRow();
       


foreach (var item in result)
{
    

}

stopwatch.Stop();

// 4. Get the elapsed time
TimeSpan elapsed = stopwatch.Elapsed;

// 5. Output the results
Console.WriteLine($"Time elapsed: {elapsed.TotalSeconds} ms");

Console.WriteLine("Done");
