using ExcelReader.src.Entity;
using System.IO.Compression;


namespace ExcelReader.src.Interfaces
{
    internal interface IReaderSheet
    {
        Task<IEnumerable<FileRowInfoExcel>> ReadSheetAsync(string sheetPartName, string[] strings, ZipArchive zipArchive);
    }
}
