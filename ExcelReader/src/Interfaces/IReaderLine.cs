using ExcelReader.src.Entity;

namespace ExcelReader.src.Interfaces
{
    internal interface IReaderLine
    {
        dynamic ReadRow(FileRowInfoExcel fileRowInfoExcel, IDictionary<string, string> columns);

        IDictionary<string, string> ReadColumns(FileRowInfoExcel fileRowInfoExcel);
    }
}
