using ExcelReader.src.Implementation;

var simple = new ReaderExcelSimple();

var result = await simple.GetNextRowAsync("Example.xlsx");

Console.WriteLine("Done");
