using ExcelReader.src.Entity;
using System.IO.Compression;

namespace ExcelReader.src.Interfaces
{
    internal interface IReaderStyles: IDisposable
    {

        void LoadInfo(ZipArchive zipArchive);
        CellStyle? GetCellStyle(int index);
    }
}
