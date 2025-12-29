using ExcelReader.src.Entity;
using System.IO.Compression;


namespace ExcelReader.src.Interfaces
{
    internal interface IReadStrings: IDisposable
    {

        public Task LoadInfoAsync(FileInfoExcel fileInfoExcel, ZipArchive zipArchive);

        public string GetStringByIndex(int index);
    }
}
