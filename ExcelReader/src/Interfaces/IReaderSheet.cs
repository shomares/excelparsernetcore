using ExcelReader.src.Entity;
using System.IO.Compression;


namespace ExcelReader.src.Interfaces
{
    internal interface IReaderSheet: IDisposable
    {
        IEnumerable<FileRowInfoExcel> ReadSheet(string sheetPartName, ZipArchive zipArchive);
    }
}
