using ExcelReader.src.Implementation;
using System.Diagnostics;
using System.Text.Json;


using var simple = new ReaderExcelSimple();

Stopwatch stopwatch = Stopwatch.StartNew();
await simple.ReadFileAsync("Example.xlsx", "sheet1", new ExcelReader.src.Config.ConfigurationReader
{
    UseMemoryForStrings = false
});


foreach (var item in simple.GetNextRow())
{

   Console.WriteLine(JsonSerializer.Serialize(item));
}

stopwatch.Stop();

// 4. Get the elapsed time
TimeSpan elapsed = stopwatch.Elapsed;

// 5. Output the results
Console.WriteLine($"Time elapsed: {elapsed.TotalSeconds} ms");

Console.WriteLine("Done");
