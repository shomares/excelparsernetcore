using ExcelReader.src.Config;

namespace ExcelReader.src.Interfaces
{
    internal interface IReadStringsFactory
    {
        IReadStrings CreateReadStrings(string fileName, ConfigurationReader? configuration = null);
    }
}
