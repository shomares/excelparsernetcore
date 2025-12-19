using ExcelReader.src.Entity;
using System.IO.Compression;


namespace ExcelReader.src.Interfaces
{
    internal interface IReadStrings: IDisposable
    {

        public IEnumerable<string> GetStrings(FileInfoExcel fileInfoExcel, ZipArchive zipArchive);
    }
}
