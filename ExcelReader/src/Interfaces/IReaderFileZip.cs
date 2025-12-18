

using ExcelReader.src.Entity;
using System.IO.Compression;

namespace ExcelReader.src.Interfaces
{
    internal interface IReaderFileZip
    {

        Task<FileInfoExcel[]> GetFileInfosAsync(ZipArchive zipArchive);
    }
}
