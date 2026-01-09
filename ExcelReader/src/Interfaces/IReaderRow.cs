using ExcelReader.src.Config;

namespace ExcelReader.src.Interfaces
{
    internal interface IReaderRow : IDisposable
    {
        IEnumerable<dynamic> GetNextRow(string sheetName);

        Task ReadFileAsync(string fileName, ConfigurationReader? configuration = null);
    }
}
