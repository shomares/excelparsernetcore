using ExcelReader.src.Config;

namespace ExcelReader.src.Interfaces
{
    internal interface IReaderRow : IDisposable
    {
        IEnumerable<dynamic> GetNextRow();

        Task ReadFileAsync(string fileName, string sheetName, ConfigurationReader? configuration = null);
    }
}
