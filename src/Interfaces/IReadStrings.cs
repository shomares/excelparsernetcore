using ExcelReader.src.Entity;
using System.IO.Compression;


namespace ExcelReader.src.Interfaces
{
    internal interface IReadStrings
    {

        public Task<string[]> ReadStringsAsync(FileInfoExcel fileInfoExcel, ZipArchive zipArchive);
    }
}
