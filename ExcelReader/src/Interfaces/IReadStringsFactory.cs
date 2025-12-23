namespace ExcelReader.src.Interfaces
{
    internal interface IReadStringsFactory
    {
        IReadStrings CreateReadStrings(string fileName);
    }
}
