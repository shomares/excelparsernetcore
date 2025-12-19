namespace ExcelReader.src.Interfaces
{
    internal interface IReaderRow: IDisposable
    {
        Task<IDictionary<string, object>> GetNextRowAsync(string file, string sheetName);
    }
}
