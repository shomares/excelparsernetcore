using ExcelReader.src.Entity;
using System.IO.Compression;


namespace ExcelReader.src.Interfaces
{
    internal interface IReaderSheet
    {
        IEnumerable<FileRowInfoExcel> ReadSheet(string sheetPartName, string[] strings, ZipArchive zipArchive);
    }
}
