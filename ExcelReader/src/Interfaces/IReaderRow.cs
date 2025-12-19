namespace ExcelReader.src.Interfaces
{
    internal interface IReaderRow
    {
        Task<IDictionary<string, object>> GetNextRowAsync(string file, string sheetName);
    }
}
